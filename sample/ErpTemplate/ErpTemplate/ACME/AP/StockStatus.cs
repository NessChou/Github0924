using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Linq;

namespace ACME
{
    public partial class StockStatus : Form
    {
        public StockStatus()
        {
            InitializeComponent();

            SetLookupCMB(cmbSize, GetSize(), false, "DataValue", "DataText", "0");
            SetLookupCMB(cmbType, GetParamsOPEN("STOCKOPEN"), false, "DataValue", "DataText", "0");
            //SetLookupCMB(cmbModel, GetParamsOPEN("STOCKOPEN"), true, "0");
            //SetLookupCMB(cmbGrade, GetParams("GRADE"), false, "DataValue", "DataText", "0");
            //SetLookupCMB(cmbVersion, GetParams("VERSION"), false, "DataValue", "DataText", "0");
            SetLookupCMB(cmbBU, GetParamsOPENBU("BU"), false, "DataValue", "DataText", "0");


        }


        private void btnSort_Click(object sender, EventArgs e)
        {
            string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
            SqlConnection MyConnection = new SqlConnection(ConnectiongString);
            //MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();

            sb.Append("SELECT P.ItemCode 產品編號, P.ItemName 品名規格, cast(P.OnHand as int) as 現有數量,");
            sb.Append(" ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0 WHERE T0.ItemCode = P.ItemCode AND LineStatus <> 'C') as int),0)  + CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 LEFT JOIN OWOR T0 ON(T0.DOCENTRY = T1.DOCENTRY) WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN('C', 'L') AND T1.ITEMCODE = P.ITEMCODE), 0) AS INT) 訂單未交量, ");
            // sb.Append(" ISNULL(datediff(d, P.lastpurdat, getdate()), 0) 庫存天數,");
            //sb.Append(" (select isnull(cast(sum(quantity) as int), 0) from por1 t0 where opencreqty > 0 and t0.itemcode = p.itemcode ) 採購未進量,");
            sb.Append("     (select isnull(cast(sum(opencreqty) as int),0) from pqt1 t0 where opencreqty >0 and t0.itemcode=p.itemcode )+    (select isnull(cast(sum(quantity) as int),0) from por1 t0 where opencreqty >0 and t0.itemcode=p.itemcode ) 採購未進量,   ");
            
            sb.Append("cast(P.OnHand as int)");
            sb.Append("-");
            sb.Append(" (ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0 WHERE T0.ItemCode = P.ItemCode AND LineStatus <> 'C') as int),0)  + CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 LEFT JOIN OWOR T0 ON(T0.DOCENTRY = T1.DOCENTRY) WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN('C', 'L') AND T1.ITEMCODE = P.ITEMCODE), 0) AS INT)) ");
            sb.Append(" - ");
            sb.Append("(select cast(t1.onhand as int)  from oitw t1   where t1.whscode = 'LB001' and t1.itemcode = p.itemcode) "); 
            sb.Append(" - "); 
            sb.Append("(SELECT cast(Sum(OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and substring(WhsCode,1,1) = 'B') ");
            sb.Append(" - ");
            sb.Append(" (SELECT cast(Sum(OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode in ('CC001', 'CC002') ) ");
            sb.Append(" - ");
            sb.Append(" (SELECT cast(Sum(OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode in ('RM001', 'RM001') ) as 可用量 ,");


            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'JC001') 借出倉, ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'ZT001') 在途倉, ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'SZ001') 深圳漢海達, ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN001') 蘇宏高倉, ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN006') 蘇偉創倉, ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN05') 深宏高倉, ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN009') '深巨航機保', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN010-1') '深巨航坪山', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN004') '廈門宏高', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN011') '武漢巨航', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'HK002') '香港宏高', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW001') '內湖倉', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW003') '平鎮倉', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW012') '聯揚倉', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW017') '新得利倉', ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW006') '經海關倉',  ");
            sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW013') '大發倉' ");

            //sb.Append(" (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'BW003') '報廢倉'");
            if (ckbOCTcon.Checked == true)
            {
                sb.Append(" ,t3.KIT,t4.itemname PartNo,ISNULL(t4.OnHand,0) T庫存量, (select isnull(cast(sum(quantity) as int),0) from AcmeSql02.DBO.por1 t0 where opencreqty >0 and t0.itemcode=t4.itemcode ) T待進貨量,ISNULL(t4.OnHand,0) + (select isnull(cast(sum(quantity) as int),0) from AcmeSql02.DBO.por1 t0 where opencreqty >0 and t0.itemcode=t4.itemcode ) T小計 ");
            }
            sb.Append(" FROM OITM P ");
            
            if (ckbOCTcon.Checked == true)
            {
                sb.Append(" left join por1 t2 on t2.itemcode = P.itemcode");
                sb.Append("  LEFT JOIN (SELECT MAX(KIT) KIT, OPENCELL FROM AcmeSqlSP.DBO.AP_OPENCELL WHERE ISNULL(KIT, '') <> ''GROUP BY SUBSTRING(KIT, 0, LEN(KIT)), OPENCELL) T3 ON(t2.itemcode = t3.opencell COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
                sb.Append("  LEFT JOIN OITM t4 on t3.KIT = t4.ItemCode COLLATE Chinese_Taiwan_Stroke_CI_AS   ");
            }
            sb.Append(" where len(P.ItemCode)= 15 And ISNULL(P.U_GROUP,'') <> 'Z&R-費用類群組' AND P.FROZENFOR = 'N' AND P.CANCELED = 'N'");


            if (ckbOCTcon.Checked == true)
            {
                if (!String.IsNullOrEmpty(txbItemcode.Text))
                {
                    char[] delims = new[] { '\r', '\n', ',' };
                    string[] NewLine = txbItemcode.Text.Split(delims, StringSplitOptions.RemoveEmptyEntries);

                    //char[] delims = new[] { '\r', '\n' };
                    //string[] NewLine = TextBox5.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string ESi in NewLine)
                    {
                        sb3.Append("'" + ESi + "',");

                    }
                    sb3.Remove(sb3.Length - 1, 1);
                    //    sb.AppendFormat(" and (P.U_TMODEL like '%{0}%') ", ddlModel.SelectedValue);
                    sb.AppendFormat(" AND   (P.ITEMCODE IN (" + sb3.ToString() + ") OR P.U_TMODEL+'.'+P.U_VERSION IN (" + sb3.ToString() + ")  OR P.U_PARTNO IN (" + sb3.ToString() + ")) ");
                    //  sb.AppendFormat(" AND   P.ITEMCODE IN (" + sb3.ToString() + ") ");
                }
                else
                {


                    if (cmbSize.SelectedValue != DBNull.Value)
                    {
                        sb.AppendFormat(" and (P.U_SIZE = '{0}') ", Convert.ToDecimal(cmbSize.SelectedValue).ToString("0.##"));
                    }

                    if (cmbType.SelectedValue != DBNull.Value)
                    {
                        sb.AppendFormat(" and (P.U_GROUP = '{0}') ", cmbType.SelectedValue.ToString());
                    }




                    if (cmbModel.SelectedValue != DBNull.Value && cmbModel.SelectedValue != null && cmbModel.SelectedValue != "none")
                    {
                        sb.AppendFormat(" and (P.U_TMODEL like '%{0}%') ", cmbModel.SelectedValue);
                    }




                    if (cmbVersion.SelectedValue != DBNull.Value && cmbVersion.SelectedValue != null)
                    {
                        sb.AppendFormat(" and (P.U_VERSION = '{0}') ", cmbVersion.SelectedValue);
                    }



                    if (cmbGrade.SelectedValue != DBNull.Value && cmbGrade.SelectedValue != null)
                    {
                        if (cmbGrade.SelectedValue.ToString().Trim() == "Z/P/N")
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}' or P.U_GRADE = '{1}' or P.U_GRADE = '{2}'or P.U_GRADE = '{3}') ", "P", "NN", "N", "Z");

                        }
                        else if (cmbGrade.SelectedValue == "P/NN")
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}' or P.U_GRADE = '{1}' or P.U_GRADE = '{2}') ", "P", "NN", "N");

                        }
                        else if (cmbGrade.SelectedValue == "Z/P")
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}' or P.U_GRADE = '{1}') ", "Z", "P");
                        }
                        else if (cmbGrade.SelectedValue == "NN")
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}' or P.U_GRADE = '{1}') ", "NN", "N");
                        }
                        else
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}') ", cmbGrade.SelectedValue);
                        }
                    }

                    if (cmbBU.SelectedValue != DBNull.Value && cmbBU.SelectedValue != null)
                    {
                        sb.AppendFormat(" and (P.U_BU LIKE '%{0}%') ", cmbBU.SelectedValue);
                    }

                    if (ckbOnHandGreatThenZero.Checked && ckbUndeliverGreaterThenZero.Checked)
                    {
                        sb.Append("      and ((ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0");
                        sb.Append("                                            WHERE T0.ItemCode=P.ItemCode AND LineStatus<>'C') as int),0)");
                        sb.Append("             +ISNULL(CAST((SELECT sum(T0.QUANTITY) FROM INV1 T0 WHERE  T0.ItemCode=P.ItemCode");
                        sb.Append("              AND BASETYPE=17 AND TRGETENTRY='') AS INT),0)");
                        sb.Append("             +CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 ");
                        sb.Append("             LEFT JOIN OWOR T0 ON (T0.DOCENTRY=T1.DOCENTRY) ");
                        sb.Append("             WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN ('C','L') ");
                        sb.Append("             AND T1.ITEMCODE=P.ITEMCODE),0) AS INT) > 0) AND  (P.OnHand > 0))");


                    }
                    else
                    {
                        if (ckbOnHandGreatThenZero.Checked)
                        {

                            sb.Append(" and P.OnHand > 0 ");

                        }
                        if (ckbUndeliverGreaterThenZero.Checked)
                        {

                            sb.Append(" and ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0");
                            sb.Append("                                WHERE T0.ItemCode=P.ItemCode AND LineStatus<>'C') as int),0)");
                            sb.Append(" +ISNULL(CAST((SELECT sum(T0.QUANTITY) FROM INV1 T0 WHERE  T0.ItemCode=P.ItemCode");
                            sb.Append("  AND BASETYPE=17 AND TRGETENTRY='') AS INT),0)");
                            sb.Append(" +CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 ");
                            sb.Append(" LEFT JOIN OWOR T0 ON (T0.DOCENTRY=T1.DOCENTRY) ");
                            sb.Append(" WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN ('C','L') ");
                            sb.Append(" AND T1.ITEMCODE=P.ITEMCODE),0) AS INT) > 0");
                        }


                    }
                    if (panel1.Text != "")
                    {
                        sb.Append("and ItemCode like '%" + panel1.Text + "%'");
                    }
                    if (txbItemcode.Text != "")
                    {
                        string[] ItemCodes = txbItemcode.Text.Split(',');
                        foreach (string itemcode in ItemCodes)
                        {
                            sb.Append("and ItemCode like '%" + itemcode + "%'");
                        }


                    }




                }
            }
            else
            {
                if (!String.IsNullOrEmpty(txbItemcode.Text))
                {
                    char[] delims = new[] { '\r', '\n', ',' };
                    string[] NewLine = txbItemcode.Text.Split(delims, StringSplitOptions.RemoveEmptyEntries);

                    //char[] delims = new[] { '\r', '\n' };
                    //string[] NewLine = TextBox5.Text.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string ESi in NewLine)
                    {
                        sb3.Append("'" + ESi + "',");

                    }
                    sb3.Remove(sb3.Length - 1, 1);
                    //    sb.AppendFormat(" and (P.U_TMODEL like '%{0}%') ", ddlModel.SelectedValue);
                    sb.AppendFormat(" AND   (P.ITEMCODE IN (" + sb3.ToString() + ") OR P.U_TMODEL+'.'+P.U_VERSION IN (" + sb3.ToString() + ")  OR P.U_PARTNO IN (" + sb3.ToString() + ")) ");
                    //  sb.AppendFormat(" AND   P.ITEMCODE IN (" + sb3.ToString() + ") ");
                }
                else
                {


                    if (cmbSize.SelectedValue != DBNull.Value)
                    {
                        sb.AppendFormat(" and (P.U_SIZE = '{0}') ", Convert.ToDecimal(cmbSize.SelectedValue).ToString("0.##"));
                    }

                    if (cmbType.SelectedValue != DBNull.Value)
                    {
                        sb.AppendFormat(" and (P.U_GROUP = '{0}') ", cmbType.SelectedValue.ToString());
                    }




                    if (cmbModel.SelectedValue != DBNull.Value && cmbModel.SelectedValue != null && cmbModel.SelectedValue != "none")
                    {
                        sb.AppendFormat(" and (P.U_TMODEL like '%{0}%') ", cmbModel.SelectedValue);
                    }




                    if (cmbVersion.SelectedValue != DBNull.Value && cmbVersion.SelectedValue != null)
                    {
                        sb.AppendFormat(" and (P.U_VERSION = '{0}') ", cmbVersion.SelectedValue);
                    }



                    if (cmbGrade.SelectedValue != DBNull.Value && cmbGrade.SelectedValue != null)
                    {
                        if (cmbGrade.SelectedValue.ToString().Trim() == "Z/P/N")
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}' or P.U_GRADE = '{1}' or P.U_GRADE = '{2}'or P.U_GRADE = '{3}') ", "P", "NN", "N", "Z");

                        }
                        else if (cmbGrade.SelectedValue == "P/NN")
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}' or P.U_GRADE = '{1}' or P.U_GRADE = '{2}') ", "P", "NN", "N");

                        }
                        else if (cmbGrade.SelectedValue == "Z/P")
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}' or P.U_GRADE = '{1}') ", "Z", "P");
                        }
                        else if (cmbGrade.SelectedValue == "NN")
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}' or P.U_GRADE = '{1}') ", "NN", "N");
                        }
                        else
                        {
                            sb.AppendFormat(" and (P.U_GRADE = '{0}') ", cmbGrade.SelectedValue);
                        }
                    }

                    if (cmbBU.SelectedValue != DBNull.Value && cmbBU.SelectedValue != null)
                    {
                        sb.AppendFormat(" and (P.U_BU LIKE '%{0}%') ", cmbBU.SelectedValue);
                    }

                    if (ckbOnHandGreatThenZero.Checked && ckbUndeliverGreaterThenZero.Checked)
                    {
                        sb.Append("      and ((ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0");
                        sb.Append("                                            WHERE T0.ItemCode=P.ItemCode AND LineStatus<>'C') as int),0)");
                        sb.Append("             +ISNULL(CAST((SELECT sum(T0.QUANTITY) FROM INV1 T0 WHERE  T0.ItemCode=P.ItemCode");
                        sb.Append("              AND BASETYPE=17 AND TRGETENTRY='') AS INT),0)");
                        sb.Append("             +CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 ");
                        sb.Append("             LEFT JOIN OWOR T0 ON (T0.DOCENTRY=T1.DOCENTRY) ");
                        sb.Append("             WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN ('C','L') ");
                        sb.Append("             AND T1.ITEMCODE=P.ITEMCODE),0) AS INT) > 0) AND  (P.OnHand > 0))");


                    }
                    else
                    {
                        if (ckbOnHandGreatThenZero.Checked)
                        {

                            sb.Append(" and P.OnHand > 0 ");

                        }
                        if (ckbUndeliverGreaterThenZero.Checked)
                        {

                            sb.Append(" and ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0");
                            sb.Append("                                WHERE T0.ItemCode=P.ItemCode AND LineStatus<>'C') as int),0)");
                            sb.Append(" +ISNULL(CAST((SELECT sum(T0.QUANTITY) FROM INV1 T0 WHERE  T0.ItemCode=P.ItemCode");
                            sb.Append("  AND BASETYPE=17 AND TRGETENTRY='') AS INT),0)");
                            sb.Append(" +CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 ");
                            sb.Append(" LEFT JOIN OWOR T0 ON (T0.DOCENTRY=T1.DOCENTRY) ");
                            sb.Append(" WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN ('C','L') ");
                            sb.Append(" AND T1.ITEMCODE=P.ITEMCODE),0) AS INT) > 0");
                        }


                    }
                    if (ckbZeroUnshow.Checked == true) 
                    {
                        sb.Append(" and (");
                        sb.Append(" cast(P.OnHand as int) > 0  ");//現有數量
                        sb.Append(" or ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0 WHERE T0.ItemCode = P.ItemCode AND LineStatus <> 'C') as int),0)  + CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 LEFT JOIN OWOR T0 ON(T0.DOCENTRY = T1.DOCENTRY) WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN('C', 'L') AND T1.ITEMCODE = P.ITEMCODE), 0) AS INT) > 0");//訂單未交量
                        sb.Append(" or (select isnull(cast(sum(opencreqty) as int),0) from pqt1 t0 where opencreqty >0 and t0.itemcode=p.itemcode )+    (select isnull(cast(sum(quantity) as int),0) from por1 t0 where opencreqty >0 and t0.itemcode=p.itemcode ) > 0 ");//採購未進量
                        sb.Append(" or cast(P.OnHand as int)- (ISNULL(cast((select sum(T0.[OpenCreQty]) AA FROM RDR1 T0 WHERE T0.ItemCode = P.ItemCode AND LineStatus <> 'C') as int),0)  + CAST(ISNULL((SELECT SUM(T1.PLANNEDQTY - T1.ISSUEDQTY) FROM WOR1 T1 LEFT JOIN OWOR T0 ON(T0.DOCENTRY = T1.DOCENTRY) WHERE T1.PLANNEDQTY > T1.ISSUEDQTY AND T0.STATUS NOT IN('C', 'L') AND T1.ITEMCODE = P.ITEMCODE), 0) AS INT))  - (select cast(t1.onhand as int)  from oitw t1   where t1.whscode = 'LB001' and t1.itemcode = p.itemcode)  - (SELECT cast(Sum(OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and substring(WhsCode,1,1) = 'B')  -  (SELECT cast(Sum(OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode in ('CC001', 'CC002') )  -  (SELECT cast(Sum(OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode in ('RM001', 'RM001') ) > 0  "); //可用量 
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'JC001') > 0 ");//借出倉
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'ZT001') > 0 ");//在途倉
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'SZ001') > 0 ");//深圳漢海達
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'SU001') > 0 ");//蘇宏高倉
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN010') > 0 ");//'深巨航機保'
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN010-1') > 0 ");//'深巨航坪山'
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN004') > 0 ");//'廈門宏高',
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'CN011') > 0 ");// '武漢巨航',
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'HK001') > 0 ");// '香港宏高'
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW001') > 0 ");//'內湖倉',
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW003') > 0 ");//'平鎮倉',
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW012') > 0 ");// '聯揚倉'
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW006') > 0 ");//'經海關倉',
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW017') > 0 ");//'新得利倉',
                        sb.Append(" or (SELECT cast((OnHand) as int) OnHand FROM OITW t0 WHERE t0.itemcode = p.itemcode and WhsCode = 'TW013') > 0 ");//大發倉'
                        sb.Append(" )");
                    }
                    if (panel1.Text != "")
                    {
                        sb.Append("and ItemCode like '%" + panel1.Text + "%'");
                    }
                    if (txbItemcode.Text != "")
                    {
                        string[] ItemCodes = txbItemcode.Text.Split(',');
                        foreach (string itemcode in ItemCodes)
                        {
                            sb.Append("and ItemCode like '%" + itemcode + "%'");
                        }


                    }




                }
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;



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

            dgvStockStatus.DataSource = ds.Tables[0];















        }

        public static void SetLookupCMB(ComboBox cmb, System.Data.DataTable dt, bool AddSelect, string Value)//combobox,資料來源,是否加提示,加在第幾項
        {
            cmb.DataSource = null;
            cmb.SelectedIndex = -1;

            cmb.DataSource = dt;
            if (AddSelect)
            {
                cmb.Items.Insert(Convert.ToInt32(Value), "Please Select");
            }
            cmb.SelectedIndex = Convert.ToInt32(Value);

        }
        public static void SetLookupCMB(ComboBox cmb, System.Data.DataTable dt, bool AddSelect, string DataValueMember, string DataDisplayMember, string Value)
        {
            cmb.DataSource = null;
            cmb.SelectedIndex = -1;

            cmb.ValueMember = DataValueMember;
            cmb.DisplayMember = DataDisplayMember;
            cmb.DataSource = dt;
            if (AddSelect)
            {
                cmb.Items.Insert(Convert.ToInt32(Value), "Please Select");
            }
            cmb.SelectedIndex = Convert.ToInt32(Value);

        }



        public static System.Data.DataTable GetParams(string KIND)
        {
            string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
            SqlConnection con = new SqlConnection(ConnectiongString);

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='" + KIND + "' order by DataValue";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "Params");
            }
            finally
            {
                con.Close();
            }
            DataRow row = ds.Tables["Params"].NewRow();
            row["DataValue"] = DBNull.Value;
            row["DataText"] = "Please Select...";
            ds.Tables["Params"].Rows.InsertAt(row, 0);

            return ds.Tables["Params"];
        }
        public static System.Data.DataTable GetSize()
        {
            string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
            SqlConnection con = new SqlConnection(ConnectiongString);


            StringBuilder sb = new StringBuilder();
            //sb.Append("           SELECT * FROM (  SELECT DISTINCT * FROM (select  distinct substring(itemcode,3,3) DataValue,");
            //sb.Append("             CASE substring(itemcode,3,1) WHEN '0' THEN CASE substring(itemcode,5,1) WHEN '0' THEN substring(itemcode,4,1) ELSE substring(itemcode,4,1)+'.'+substring(itemcode,5,1) ");
            //sb.Append("             END ELSE CASE substring(itemcode,5,1) WHEN '0' THEN substring(itemcode,3,2) ELSE substring(itemcode,3,2)+'.'+substring(itemcode,5,1)  END END DataText from oitm  where substring(itemcode,1,1) IN ('T','J')  ");
            //sb.Append("             AND substring(itemcode,3,1) BETWEEN '0' AND '9'   AND FROZENFOR = 'N'    ");
            //sb.Append("           UNION ALL");
            //sb.Append("             select  distinct substring(itemcode,2,3) DataValue,");
            //sb.Append("             CASE substring(itemcode,2,1) WHEN '0' THEN CASE substring(itemcode,4,1) WHEN '0' THEN substring(itemcode,3,1) ELSE substring(itemcode,3,1)+'.'+substring(itemcode,4,1) ");
            //sb.Append("             END ELSE CASE substring(itemcode,4,1) WHEN '0' THEN substring(itemcode,2,2) ELSE substring(itemcode,2,2)+'.'+substring(itemcode,4,1)  END END DataText from oitm  where");
            //sb.Append("          SUBSTRING(ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            //sb.Append("                  SUBSTRING(ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            //sb.Append("                  SUBSTRING(ITEMCODE,3,1) LIKE '[0-9]%'");
            //sb.Append("                 AND SUBSTRING(ITEMCODE,4,1) LIKE '[0-9]%'   AND FROZENFOR = 'N'    ) AS A");
            //sb.Append("      ) AS A WHERE ISNUMERIC(DataValue)  <> 0");
            //sb.Append("     order by DataValue ");

            sb.Append("    SELECT DISTINCT CAST(U_SIZE AS DECIMAL(5,2)) DataValue,U_SIZE DataText FROM OITM ");
            sb.Append("  WHERE FROZENFOR = 'N'  AND U_SIZE <> 'X' ORDER BY CAST(U_SIZE AS DECIMAL(5,2))");


            SqlDataAdapter da = new SqlDataAdapter(sb.ToString(), con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "Params");
            }
            finally
            {
                con.Close();
            }
            try
            {
                DataRow row = ds.Tables["Params"].NewRow();
                row["DataValue"] = DBNull.Value;
                row["DataText"] = "Please Select...";
                ds.Tables["Params"].Rows.InsertAt(row, 0);
            }
            catch (Exception ex)
            {

            }


            return ds.Tables["Params"];
        }
        public static System.Data.DataTable GetParamsSP(string KIND)
        {
            string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
            SqlConnection con = new SqlConnection(ConnectiongString);

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='" + KIND + "' order by DataValue";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "Params");
            }
            finally
            {
                con.Close();
            }

            DataRow row = ds.Tables["Params"].NewRow();
            row["DataValue"] = DBNull.Value;
            row["DataText"] = "Please Select...";
            ds.Tables["Params"].Rows.InsertAt(row, 0);

            return ds.Tables["Params"];
        }

        public static System.Data.DataTable GetParamsOPEN(string KIND)
        {
            string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
            SqlConnection con = new SqlConnection(ConnectiongString);

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='" + KIND + "' order by [ID]";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "Params");
            }
            finally
            {
                con.Close();
            }
            DataRow row = ds.Tables["Params"].NewRow();
            row["DataValue"] = DBNull.Value;
            row["DataText"] = "Please Select...";
            ds.Tables["Params"].Rows.InsertAt(row, 0);

            return ds.Tables["Params"];
        }
        public static System.Data.DataTable GetParamsOPENBU(string KIND)
        {
            string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesqlsp";
            SqlConnection con = new SqlConnection(ConnectiongString);

            string sql = "SELECT PARAM_NO as DataValue,PARAM_DESC as DataText FROM RMA_PARAMS where param_kind='" + KIND + "' AND (ID = '186' OR ID = '187' OR ID = '204' OR ID = '206') order by [ID]";

            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "Params");
            }
            finally
            {
                con.Close();
            }
            DataRow row = ds.Tables["Params"].NewRow();
            row["DataValue"] = DBNull.Value;
            row["DataText"] = "Please Select...";
            ds.Tables["Params"].Rows.InsertAt(row, 0);

            return ds.Tables["Params"];
        }
        private System.Data.DataTable MakeTableMODEL()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("DataValue", typeof(string));
            dt.Columns.Add("DataText", typeof(string));

            DataColumn[] colPk = new DataColumn[1];
            colPk[0] = dt.Columns["DataValue"];
            dt.PrimaryKey = colPk;
            return dt;
        }
        public System.Data.DataTable GetmodelINF(string KIND, string GROUP)
        {
            string ConnectiongString = "server=acmesap;pwd=@rmas;uid=sapdbo;database=acmesql02";
            SqlConnection con = new SqlConnection(ConnectiongString);


            StringBuilder sb = new StringBuilder();

            sb.Append(" select  distinct U_TMODEL DataValue from oitm  ");
            sb.Append(" where ISNULL(U_GROUP,'') <> 'Z&R-費用類群組' AND FROZENFOR = 'N' AND U_SIZE='" + KIND + "'");
            if (cmbType.SelectedValue != DBNull.Value && cmbType.SelectedValue != null)
            {
                sb.Append("  AND U_GROUP='" + GROUP + "' ");
            }

            SqlDataAdapter da = new SqlDataAdapter(sb.ToString(), con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "Params");
            }
            catch (Exception ex)
            {

            }
            finally
            {
                con.Close();
            }
            return ds.Tables["Params"];
        }

        private void cmbSize_TextChanged(object sender, EventArgs e)
        {
            if (cmbSize.Text != "Please Select...")
            {
                System.Data.DataTable k1 = new System.Data.DataTable();
                if (cmbSize.SelectedValue != null)
                {
                    k1 = GetmodelINF(Convert.ToDecimal(cmbSize.SelectedValue).ToString("0.##"), cmbType.SelectedValue.ToString());
                }
                else 
                {
                    k1 = GetmodelINF((cmbSize.SelectedValue).ToString(), cmbType.SelectedValue.ToString());
                }
               
                System.Data.DataTable dtCost = MakeTableMODEL();
                DataRow dr = null;
                DataRow row2;

                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    string MODEL = k1.Rows[i][0].ToString();
                    string[] arrurl = MODEL.Split(new Char[] { '/' });
                    foreach (string i2 in arrurl)
                    {
                        row2 = dtCost.Rows.Find(i2);
                        if (row2 == null)
                        {
                            dr = dtCost.NewRow();
                            dr["DataValue"] = i2;
                            dr["DataText"] = i2;
                            dtCost.Rows.Add(dr);
                        }
                    }

                }
                DataRow row = dtCost.NewRow();
                row["DataValue"] = "none";
                row["DataText"] = "Please Select...";
                dtCost.Rows.InsertAt(row, 0);
                SetLookupCMB(cmbModel, dtCost, false, "DataValue", "DataText", "0");

            }

        }

        private void cmbType_TextChanged(object sender, EventArgs e)
        {
            if (cmbType.Text != "Please Select...")
            {
                System.Data.DataTable k1 = new System.Data.DataTable();
                if (cmbSize.SelectedValue != null)
                {
                    k1 = GetmodelINF(Convert.ToDecimal(cmbSize.SelectedValue).ToString("0.##"), cmbType.SelectedValue.ToString());
                }
                else 
                {
                    k1 = GetmodelINF((cmbSize.SelectedValue).ToString(), cmbType.SelectedValue.ToString());
                }
                

                System.Data.DataTable dtCost = MakeTableMODEL();
                DataRow dr = null;
                DataRow row2;

                for (int i = 0; i <= k1.Rows.Count - 1; i++)
                {
                    string MODEL = k1.Rows[i][0].ToString();
                    string[] arrurl = MODEL.Split(new Char[] { '/' });
                    foreach (string i2 in arrurl)
                    {
                        row2 = dtCost.Rows.Find(i2);
                        if (row2 == null)
                        {
                            dr = dtCost.NewRow();
                            dr["DataValue"] = i2;
                            dr["DataText"] = i2;
                            dtCost.Rows.Add(dr);
                        }
                    }

                }
                DataRow row = dtCost.NewRow();
                row["DataValue"] = "none";
                row["DataText"] = "Please Select...";
                dtCost.Rows.InsertAt(row, 0);
                SetLookupCMB(cmbModel, dtCost, false, "DataValue", "DataText", "0");

            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            System.Data.DataTable table = (System.Data.DataTable)(dgvStockStatus.DataSource) ;
            string date = DateTime.Now.ToString("yyyy/MM");
            string FileNameTemplate = GetExePath() + "\\Excel\\wh\\貨況表.xls";
            string FileName = GetExePath() + "\\Excel\\temp\\" + date + "貨況表.xls";
            //System.Data.DataTable dtData = GetDataSort(table, "notZero");//把一樣的相加 小計不為零
            WriteDataTableToExcel(table, FileNameTemplate, FileName);
        }
        private System.Data.DataTable GetDataSort(System.Data.DataTable dtData, string flag)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            DataRow rw;
            dt = MakeTable();

            foreach (DataRow row in dtData.Rows)
            {
                if (flag == "notZero" && row["小計"].ToString() != "0" && row["小計"].ToString() != "")
                {
                    //避免重複加入
                    int onhand = 0;//庫存量
                    int wait = 0;//待進貨量
                    int sum = 0;//小計
                    int Tonhand = 0;//T庫存量
                    int Twait = 0;//T待進貨量
                    int Tsum = 0;//T小計
                    string test = row["KIT"].ToString();
                    if (row["KIT"].ToString() != "" && row["KIT"].ToString() != null)
                    {
                        if (dt.Select("產品編號 = '" + row["產品編號"].ToString() + "' and KIT LIKE  '" + row["KIT"].ToString().Substring(0, 14) + "*'").Length != 0)
                        {
                            continue;
                        }
                        //DataRow[] rws = dtData.Select("產品編號 = '" + row["產品編號"].ToString() + "' and KIT like '" + row["KIT"].ToString().Substring(0,14) + "*'");


                        //var rws = dtData.AsEnumerable().Where(r => r.Field<string>("KIT").Contains(row["KIT"].ToString().Substring(0, 14)) && r.Field<string>("產品編號") == row["產品編號"].ToString());

                        var rws = from dr in dtData.AsEnumerable()
                                  where dr.Field<string>("產品編號") == row["產品編號"].ToString() && dr.Field<string>("KIT") != null && dr.Field<string>("KIT").Substring(0, 14) == row["KIT"].ToString().Substring(0, 14)
                                  select dr;

                        int rowcount = rws.Count();
                        //int rowcount = rws.Length;
                        foreach (DataRow rwss in rws)
                        {
                            onhand += Convert.ToInt32(rwss["庫存量"]);
                            wait += Convert.ToInt32(rwss["待進貨量"]);
                            sum += Convert.ToInt32(rwss["小計"]);
                            Tonhand = Convert.ToInt32(rwss["T庫存量"]);
                            Twait = Convert.ToInt32(rwss["T待進貨量"]);
                            Tsum = Convert.ToInt32(rwss["T小計"]);
                        }

                    }
                    else
                    {
                        DataRow[] rws = dtData.Select("產品編號 = '" + row["產品編號"].ToString() + "' and KIT = null");
                        onhand += Convert.ToInt32(row["庫存量"]);
                        wait += Convert.ToInt32(row["待進貨量"]);
                        sum += Convert.ToInt32(row["小計"]);
                        Tonhand = Convert.ToInt32(row["T庫存量"]);
                        Twait = Convert.ToInt32(row["T待進貨量"]);
                        Tsum = Convert.ToInt32(row["T小計"]);


                    }

                    rw = dt.NewRow();
                    rw["產品編號"] = row["產品編號"].ToString();
                    rw["KIT"] = row["KIT"].ToString();
                    rw["PartNo"] = row["PartNo"].ToString();
                    rw["庫存量"] = onhand;
                    rw["待進貨量"] = wait;
                    rw["小計"] = sum;
                    rw["T庫存量"] = Tonhand;
                    rw["T待進貨量"] = Twait;
                    rw["T小計"] = Tsum;
                    dt.Rows.Add(rw);
                }
                else if (flag == "Zero" && row["小計"].ToString() == "0" && row["小計"].ToString() != "")
                {
                    //避免重複加入
                    int onhand = 0;//庫存量
                    int wait = 0;//待進貨量
                    int sum = 0;//小計
                    int Tonhand = 0;//T庫存量
                    int Twait = 0;//T待進貨量
                    int Tsum = 0;//T小計
                    DataRow[] rws = dtData.Select("產品編號 = '" + row["產品編號"].ToString() + "' and KIT = '" + row["KIT"].ToString() + "'");
                    for (int i = 0; i < rws.Length; i++)
                    {
                        onhand = Convert.ToInt32(rws[i]["庫存量"]);
                        wait = Convert.ToInt32(rws[i]["待進貨量"]);
                        sum = Convert.ToInt32(rws[i]["小計"]);
                        Tonhand = Convert.ToInt32(rws[i]["T庫存量"]);
                        Twait = Convert.ToInt32(rws[i]["T待進貨量"]);
                        Tsum = Convert.ToInt32(rws[i]["T小計"]);
                    }
                    rw = dt.NewRow();
                    rw["產品編號"] = row["產品編號"].ToString();
                    rw["KIT"] = row["KIT"].ToString();
                    rw["PartNo"] = row["PartNo"].ToString();
                    rw["庫存量"] = row["庫存量"].ToString();
                    rw["待進貨量"] = row["待進貨量"].ToString();
                    rw["小計"] = row["小計"].ToString();
                    rw["T庫存量"] = 0;
                    rw["T待進貨量"] = 0;
                    rw["T小計"] = 0;
                    dt.Rows.Add(rw);
                }
            }


            return dt;
        }
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }
        private void WriteDataTableToExcel(System.Data.DataTable dt, string DirTemplate, string Dir)
        {
            Microsoft.Office.Interop.Excel.Application excel = null;
            Microsoft.Office.Interop.Excel.Workbook excelworkBook = null;
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            // Microsoft.Office.Interop.Excel.Range excelCellrange;
            object oMissing = System.Reflection.Missing.Value;


            //  get Application object.
            excel = new Microsoft.Office.Interop.Excel.Application();
            excel.Visible = true;
            excel.DisplayAlerts = false;



            //Interop params
            string ItemCode = "";
            string ItemCodeAll = "";
            foreach (DataRow itemcode in dt.Rows)
            {
                string ItemCodetmp = Convert.ToString(itemcode["產品編號"]).Split('.')[0] + "." + Convert.ToString(itemcode["產品編號"]).Split('.')[1].Substring(1, 1);//ex G170ETN01.00022 => G170ETN01.0 小數點前加小數點後第二位
                if (!ItemCode.Contains(ItemCodetmp))
                {
                    ItemCode += ItemCodetmp + ",";
                    ItemCodeAll += Convert.ToString(itemcode["產品編號"]) + ",";//完整的字串 之後做sort用
                }

            }
            ItemCode = ItemCode.Substring(0, ItemCode.Length - 1);
            int ItemCodeCount = ItemCode.Split(',').Length;

            try
            {


                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);
                SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.ActiveSheet;

                // Workk sheet
                SheetTemplate.Copy(Type.Missing, excelworkBook.Sheets[excelworkBook.Sheets.Count]);
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets[excelworkBook.Sheets.Count];

                excelSheet.Name = "貨況表";
                if (ckbOCTcon.Checked == true)
                {
                    WriteDataTableToSheetByArray(dt, excelSheet, true);
                }
                else
                {
                    WriteDataTableToSheetByArray(dt, excelSheet, false);
                }





                //now save the workbook and exit Excel
                //excelworkBook.SaveAs(saveAsLocation);
                excelworkBook.SaveAs(Dir, XlFileFormat.xlWorkbookNormal,
                      "", "", Type.Missing, Type.Missing,
                    XlSaveAsAccessMode.xlNoChange,
                    1, false, Type.Missing, Type.Missing, Type.Missing);



                SheetTemplate.Delete();
                excelworkBook.Close();
                System.Diagnostics.Process.Start(Dir);

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();

                //System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelworkBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);



                excelSheet = null;
                // excelCellrange = null;
                excelworkBook = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }
        }
        private static void WriteDataTableToSheetByArray(System.Data.DataTable dataTable,
            Worksheet worksheet,bool OpenCellTag)
        {
            Microsoft.Office.Interop.Excel.Range excelRange;
            int rows = dataTable.Rows.Count + 1;
            int columns = dataTable.Columns.Count;
            int rownow = 1;


            var data = new object[rows, columns];

            int rowcount = 0;
            for (int i = 1; i <= columns; i++)
            {
                data[rowcount, i - 1] = dataTable.Columns[i - 1].ColumnName;
            }

            rowcount += 1;
            foreach (DataRow datarow in dataTable.Rows)
            {
                for (int i = 1; i <= dataTable.Columns.Count; i++)
                {

                    // Filling the excel file 
                    data[rowcount, i - 1] = datarow[i - 1].ToString();


                }


                rowcount += 1;
            }

            var startCell = (Range)worksheet.Cells[1, 1];
            var endCell = (Range)worksheet.Cells[rows, columns];
            var writeRange = worksheet.Range[startCell, endCell];

            //aRange.Columns.AutoFit();

            writeRange.Value2 = data;
            rowcount = 2;//第二行開始
            /*
            if (OpenCellTag == true) 
            {
                //KIT相同合併
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    try
                    {
                        DataRow[] rowss = dataTable.Select("KIT = '" + dataTable.Rows[i]["KIT"].ToString() + "' and PartNo = '" + dataTable.Rows[i]["PartNo"].ToString() + "'");
                        if (rowss.Length > 1 && dataTable.Rows[i]["KIT"].ToString() != "")
                        {
                            int j = rowss.Length - 1;
                            worksheet.get_Range("E" + (i + 2).ToString(), "E" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("E" + (i + 2).ToString(), "E" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("F" + (i + 2).ToString(), "F" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("F" + (i + 2).ToString(), "F" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("G" + (i + 2).ToString(), "G" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("G" + (i + 2).ToString(), "G" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("H" + (i + 2).ToString(), "H" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("H" + (i + 2).ToString(), "H" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("I" + (i + 2).ToString(), "I" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("I" + (i + 2).ToString(), "I" + (i + 2 + j).ToString()).MergeCells);
                        }



                        rowcount += rowss.Length;
                        i = rowss.Length > 1 ? i += rowss.Length - 1 : i;
                    }
                    catch (Exception ex)
                    {

                    }
                }
                //產品編號相同合併
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    try
                    {
                        DataRow[] rowss = dataTable.Select("產品編號 = '" + dataTable.Rows[i]["產品編號"].ToString() + "' and 庫存量 = '" + dataTable.Rows[i]["庫存量"].ToString() + "'and 待進貨量 = '" + dataTable.Rows[i]["待進貨量"].ToString() + "'");
                        if (rowss.Length > 1 && dataTable.Rows[i]["KIT"].ToString() != "")
                        {
                            int j = rowss.Length - 1;
                            worksheet.get_Range("A" + (i + 2).ToString(), "A" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("A" + (i + 2).ToString(), "A" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("B" + (i + 2).ToString(), "B" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("B" + (i + 2).ToString(), "B" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("C" + (i + 2).ToString(), "C" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("C" + (i + 2).ToString(), "C" + (i + 2 + j).ToString()).MergeCells);
                            worksheet.get_Range("D" + (i + 2).ToString(), "D" + (i + 2 + j).ToString()).Merge(worksheet.get_Range("D" + (i + 2).ToString(), "D" + (i + 2 + j).ToString()).MergeCells);

                        }



                        rowcount += rowss.Length;
                        i = rowss.Length > 1 ? i += rowss.Length - 1 : i;
                    }
                    catch (Exception ex)
                    {

                    }
                }

            }
            */
            writeRange.Columns.AutoFit();
        }
        private System.Data.DataTable MakeTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();


            dt.Columns.Add("產品編號", typeof(string));

            dt.Columns.Add("庫存量", typeof(int));


            dt.Columns.Add("待進貨量", typeof(int));

            dt.Columns.Add("小計", typeof(int));

            dt.Columns.Add("KIT", typeof(string));
            //TCON
            dt.Columns.Add("PartNo", typeof(string));

            dt.Columns.Add("T庫存量", typeof(int));
            //TCON
            dt.Columns.Add("T待進貨量", typeof(int));
            //TCON
            dt.Columns.Add("T小計", typeof(int));


            dt.TableName = "dt";

            //寫入資料
            //DataRow dr;
            //dr = dt.NewRow();
            //dr["Item"] = "訂單張數";
            //dt.Rows.Add(dr);


            return dt;
        }

        private void cmbSize_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}

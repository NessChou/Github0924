using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace ACME
{
    public partial class StockCheck : Form
    {
        public StockCheck()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            ViewBatchPayment();
        }
        private void ViewBatchPayment()
        {
 
            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                     select distinct Convert(varchar(10),docdate,112)  日期, T0.itemcode 料號,case t0.transtype ");
            sb.Append("                          when 13 then 'AR' when 14 then 'AR貸項' ");
            sb.Append("                          when 15 then '交貨' when 16 then '銷售退貨' ");
            sb.Append("                          when 18 then 'AP' when 19 then 'AP貸項' ");
            sb.Append("                          when 20 then '收貨採購單' when 59 then '收貨單' ");
            sb.Append("                          when 60 then '發貨單' when 67 then '庫存調撥' ");
            sb.Append("                          ELSE CAST(t0.transtype AS NVARCHAR) END  總類");
            sb.Append("                          ,t0.base_ref 單號,t0.doclinenum 列號,cardname 客戶名稱,");
            sb.Append("                         case CAST(CAST(isnull(TT1.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T1.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else  CAST(CAST(isnull(TT1.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T1.QTY,0) AS INT) AS NVARCHAR) end '報廢倉-內湖',");
            sb.Append("                         case CAST(CAST(isnull(TT2.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T2.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT2.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T2.QTY,0) AS INT) AS NVARCHAR) end '報廢倉-友福',        ");
            sb.Append("                         case CAST(CAST(isnull(TT3.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T3.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT3.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T3.QTY,0) AS INT) AS NVARCHAR) end'蘇州倉-宏高',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT4.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T4.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT4.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T4.QTY,0) AS INT) AS NVARCHAR) end '上海倉-瀚運',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT5.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T5.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT5.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T5.QTY,0) AS INT) AS NVARCHAR) end '深圳倉',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT6.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T6.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT6.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T6.QTY,0) AS INT) AS NVARCHAR) end '廈門倉-宏高',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT7.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T7.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT7.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T7.QTY,0) AS INT) AS NVARCHAR) end '深圳-宏高倉',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT8.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T8.QTY,0)AS INT) AS NVARCHAR)  when '0/0' then '' else CAST(CAST(isnull(TT8.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T8.QTY,0) AS INT) AS NVARCHAR) end '香港倉-軟通',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT9.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T9.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT9.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T9.QTY,0) AS INT) AS NVARCHAR) end '香港倉-宏高',");
            sb.Append("                         case CAST(CAST(isnull(TT10.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T10.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT10.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T10.QTY,0) AS INT) AS NVARCHAR) end  '香港倉-浩洋',        ");
            sb.Append("                         case CAST(CAST(isnull(TT11.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T11.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT11.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T11.QTY,0) AS INT) AS NVARCHAR) end  '借出倉',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT12.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T12.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT12.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T12.QTY,0) AS INT) AS NVARCHAR) end  '借入倉',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT13.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T13.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT13.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T13.QTY,0) AS INT) AS NVARCHAR) end  '在途倉',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT14.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T14.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT14.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T14.QTY,0) AS INT) AS NVARCHAR) end  '維修倉-內湖',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT15.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T15.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT15.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T15.QTY,0) AS INT) AS NVARCHAR) end  '維修倉-友福',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT16.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T16.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT16.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T16.QTY,0) AS INT) AS NVARCHAR) end  '內湖倉',   ");
            sb.Append("                         case CAST(CAST(isnull(TT17.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T17.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT17.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T17.QTY,0) AS INT) AS NVARCHAR) end  '友福倉',               ");
            sb.Append("                         case CAST(CAST(isnull(TT18.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T18.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT18.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T18.QTY,0) AS INT) AS NVARCHAR) end  '平鎮倉',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT19.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T19.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT19.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T19.QTY,0) AS INT) AS NVARCHAR) end  '友達倉-國內直送',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT20.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T20.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT20.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T20.QTY,0) AS INT) AS NVARCHAR) end  '友達倉-合作外銷',                   ");
            sb.Append("                         case CAST(CAST(isnull(TT21.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T21.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT21.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T21.QTY,0) AS INT) AS NVARCHAR) end  '經海關倉',   ");
            sb.Append("                         case CAST(CAST(isnull(TT22.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T22.QTY,0) AS INT) AS NVARCHAR) when '0/0' then '' else CAST(CAST(isnull(TT22.數量,0) AS INT)AS NVARCHAR)+'/'+CAST(CAST(isnull(T22.QTY,0) AS INT) AS NVARCHAR) end  '非經海關倉'              ");
            sb.Append("                          from oinm t0  INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T0.ItemCode ");
            sb.Append("                          LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='BW001'  ) T1 ON (T0.transtype=T1.transtype and T0.base_ref=T1.base_ref and T0.doclinenum=T1.doclinenum )");
            sb.Append("                        LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='BW001'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT1 ON (T0.ITEMCODE=TT1.ITEMCODE)");
            sb.Append("                          LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='BW002' ) T2 ON (T0.transtype=T2.transtype and T0.base_ref=T2.base_ref and T0.doclinenum=T2.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='BW002'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT2 ON (T0.ITEMCODE=TT2.ITEMCODE)           ");
            sb.Append("               LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='CN001' ) T3 ON (T0.transtype=T3.transtype and T0.base_ref=T3.base_ref and T0.doclinenum=T3.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='CN001'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT3 ON (T0.ITEMCODE=TT3.ITEMCODE)   ");
            sb.Append("                       LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='CN002' ) T4 ON (T0.transtype=T4.transtype and T0.base_ref=T4.base_ref and T0.doclinenum=T4.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='CN002'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT4 ON (T0.ITEMCODE=TT4.ITEMCODE)             ");
            sb.Append("                LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='CN003' ) T5 ON (T0.transtype=T5.transtype and T0.base_ref=T5.base_ref and T0.doclinenum=T5.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='CN003'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT5 ON (T0.ITEMCODE=TT5.ITEMCODE)               ");
            sb.Append("              LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='CN004' ) T6 ON (T0.transtype=T6.transtype and T0.base_ref=T6.base_ref and T0.doclinenum=T6.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='CN004'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT6 ON (T0.ITEMCODE=TT6.ITEMCODE)               ");
            sb.Append("              LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='CN05' ) T7 ON (T0.transtype=T7.transtype and T0.base_ref=T7.base_ref and T0.doclinenum=T7.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='CN05'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT7 ON (T0.ITEMCODE=TT7.ITEMCODE)             ");
            sb.Append("                LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='HK001' ) T8 ON (T0.transtype=T8.transtype and T0.base_ref=T8.base_ref and T0.doclinenum=T8.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='HK001'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT8 ON (T0.ITEMCODE=TT8.ITEMCODE)              ");
            sb.Append("               LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='HK002' ) T9 ON (T0.transtype=T9.transtype and T0.base_ref=T9.base_ref and T0.doclinenum=T9.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='HK002'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT9 ON (T0.ITEMCODE=TT9.ITEMCODE)                ");
            sb.Append("             LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='HK003' ) T10 ON (T0.transtype=T10.transtype and T0.base_ref=T10.base_ref and T0.doclinenum=T10.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='HK003'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT10 ON (T0.ITEMCODE=TT10.ITEMCODE)               ");
            sb.Append("              LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='LB001' ) T11 ON (T0.transtype=T11.transtype and T0.base_ref=T11.base_ref and T0.doclinenum=T11.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='LB001'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT11 ON (T0.ITEMCODE=TT11.ITEMCODE)                ");
            sb.Append("             LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='LB002' ) T12 ON (T0.transtype=T12.transtype and T0.base_ref=T12.base_ref and T0.doclinenum=T12.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='LB002'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT12 ON (T0.ITEMCODE=TT12.ITEMCODE)                ");
            sb.Append("             LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='OT001' ) T13 ON (T0.transtype=T13.transtype and T0.base_ref=T13.base_ref and T0.doclinenum=T13.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='OT001'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT13 ON (T0.ITEMCODE=TT13.ITEMCODE)   ");
            sb.Append("                          LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='RM001') T14 ON (T0.transtype=T14.transtype and T0.base_ref=T14.base_ref and T0.doclinenum=T14.doclinenum )");
            sb.Append("                LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='RM001'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT14 ON (T0.ITEMCODE=TT14.ITEMCODE)             ");
            sb.Append("              LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='RM002') T15 ON (T0.transtype=T15.transtype and T0.base_ref=T15.base_ref and T0.doclinenum=T15.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='RM002'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT15 ON (T0.ITEMCODE=TT15.ITEMCODE)                ");
            sb.Append("             LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='TW001' ) T16 ON (T0.transtype=T16.transtype and T0.base_ref=T16.base_ref and T0.doclinenum=T16.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='TW001'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT16 ON (T0.ITEMCODE=TT16.ITEMCODE)             ");
            sb.Append("                LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='TW002' ) T17 ON (T0.transtype=T17.transtype and T0.base_ref=T17.base_ref and T0.doclinenum=T17.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='TW002'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT17 ON (T0.ITEMCODE=TT17.ITEMCODE)               ");
            sb.Append("              LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='TW003' ) T18 ON (T0.transtype=T18.transtype and T0.base_ref=T18.base_ref and T0.doclinenum=T18.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='TW003'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT18 ON (T0.ITEMCODE=TT18.ITEMCODE)       ");
            sb.Append("                      LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum  FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='TW004' ) T19 ON (T0.transtype=T19.transtype and T0.base_ref=T19.base_ref and T0.doclinenum=T19.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='TW004'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT19 ON (T0.ITEMCODE=TT19.ITEMCODE)              ");
            sb.Append("             LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='TW005' ) T20 ON (T0.transtype=T20.transtype and T0.base_ref=T20.base_ref and T0.doclinenum=T20.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='TW005'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT20 ON (T0.ITEMCODE=TT20.ITEMCODE)             ");
            sb.Append("              LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='TW006' ) T21 ON (T0.transtype=T21.transtype and T0.base_ref=T21.base_ref and T0.doclinenum=T21.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='TW006'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT21 ON (T0.ITEMCODE=TT21.ITEMCODE) ");
            sb.Append("                          LEFT JOIN (SELECT CASE (INQTY) WHEN 0 THEN (OUTQTY)*-1 ELSE (INQTY) END QTY,base_ref,transtype,doclinenum  FROM OINM");
            sb.Append("                           WHERE WAREHOUSE='TW007' ) T22 ON (T0.transtype=T22.transtype and T0.base_ref=T22.base_ref and T0.doclinenum=T22.doclinenum )");
            sb.Append("              LEFT JOIN (SELECT SUM(T0.[InQty])-SUM(T0.[OutQty]) 數量,T0.ITEMCODE FROM OINM T0 WHERE WAREHOUSE='TW007'");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) < @startday GROUP BY ITEMCODE  ) TT22 ON (T0.ITEMCODE=TT22.ITEMCODE)             ");
            sb.Append("              where  ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組' ");
            sb.Append("              and (t0.OUTQTY <> 0 or t0.INQTY <> 0) ");
            sb.Append("             AND Convert(varchar(10),T0.DOCDATE,112) between @startday and @endday order by t0.itemcode,Convert(varchar(10),docdate,112)");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@startday", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@endday ", textBox2.Text));

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

            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }

        private void StockCheck_Load(object sender, EventArgs e)
        {
            textBox1.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox2.Text = DateTime.Now.ToString("yyyyMMdd"); 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);

        }

    }
}
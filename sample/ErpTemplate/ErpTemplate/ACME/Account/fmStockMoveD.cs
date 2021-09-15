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
    public partial class fmStockMoveD : Form
    {
        public string DATETIME1;
        public string DATETIME2;
        public string ITEMCODE;
        public string DOCTYPE;
        public fmStockMoveD()
        {
            InitializeComponent();
        }
        private System.Data.DataTable GetItemHisListByTransType(string DocDate1, string DocDate2, string ITEMCODE)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();



            sb.Append("              SELECT DBO.fun_SAPDOC(TRANSTYPE) 單據總類,Convert(varchar(10),T0.[DocDate],111) 日期,BASE_REF 單號,T0.CARDNAME 客戶名稱,(T0.[InQty] - T0.[OutQty]) 數量,(T0.[TransValue]) 金額,comments 備註");
            sb.Append("              FROM  [dbo].[OINM] T0   ");
            sb.Append("              INNER  JOIN [dbo].[OITM] T1  ON  T1.[ItemCode] = T0.ItemCode   ");
            sb.Append("              WHERE ISNULL(T1.U_GROUP,'') <> 'Z&R-費用類群組'  AND ((T0.[InQty] - T0.[OutQty])+(T0.[TransValue])) <> 0  ");
            sb.Append("and  T0.[DocDate] >= @DocDate1 ");
            sb.Append("And    T0.[DocDate] <= @DocDate2 ");
            if (DOCTYPE == "本期進貨")
            {
                sb.Append("and  TRANSTYPE IN (18,20)");
            }
            if (DOCTYPE == "進貨退出折讓")
            {
                sb.Append("and  TRANSTYPE IN (19,21)");
            }
            if (DOCTYPE == "本期銷貨")
            {
                sb.Append("and  TRANSTYPE IN (13,15)");
            }
            if (DOCTYPE == "銷貨退回")
            {
                sb.Append("and  TRANSTYPE IN (14,16)");
            }
            if (DOCTYPE == "本期調整")
            {
                sb.Append("and  TRANSTYPE IN (59,60)");
            }
            if (DOCTYPE == "本期調撥")
            {
                sb.Append("and  TRANSTYPE IN (67)");
            }
            sb.Append("And    T0.ITEMCODE = @ITEMCODE order by  T0.[DocDate]");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DocDate1", DocDate1));
            command.Parameters.Add(new SqlParameter("@DocDate2", DocDate2));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "PRODUCT");
            }
            finally
            {
                connection.Close();
            }

            System.Data.DataTable dt = ds.Tables["PRODUCT"];

            return dt;

        }

        private void fmStockMoveD_Load(object sender, EventArgs e)
        {
            System.Data.DataTable G1 = GetItemHisListByTransType(DATETIME1, DATETIME2, ITEMCODE);
            dataGridView1.DataSource = G1;

            for (int i = 4; i <= 5; i++)
            {
                DataGridViewColumn col = dataGridView1.Columns[i];
   
                col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

                col.DefaultCellStyle.Format = "#,##0";
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;
namespace ACME
{
    public partial class ADGETPAY : Form
    {
        System.Data.DataTable dtAD = null;
        string str16 = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public ADGETPAY()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {

            System.Data.DataTable F1= GetTYPES("TFT");

  

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\AD\\宇豐應收.xlsx";
            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //產生 Excel ReportdataGridView1
            ExcelReport.FIONAT(GetTYPES("TFT"), GetTYPES("PV"), ExcelTemplate, OutPutFile, GetMenu.DayS(textBox1.Text));
        }

        public System.Data.DataTable GetTYPES(string DTYPE)
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();

          

            sb.Append(" SELECT 客戶編號,客戶名稱,SUM(金額NTD) '金額NTD',SUM([<0日]) '<0日',SUM([30日以下]) '30日以下', SUM([60日～31日]) '60日～31日'");
            sb.Append(" ,SUM([90日～61日]) '90日～61日',SUM([180日～91日]) '180日～91日',SUM([181日以上]) '181日以上' FROM (SELECT 客戶編號,客戶名稱,CAST(結餘 AS INT) 金額NTD,  ");
            sb.Append(" ISNULL(CASE WHEN 逾期天數 <0 THEN CAST(結餘 AS INT) END,'') '<0日', ");
            sb.Append(" ISNULL(CASE WHEN 逾期天數 BETWEEN 0 AND 30 THEN CAST(結餘 AS INT) END,'') '30日以下', ");
            sb.Append(" ISNULL(CASE WHEN 逾期天數 BETWEEN 31 AND 60 THEN CAST(結餘 AS INT) END,'') '60日～31日', ");
            sb.Append(" ISNULL(CASE WHEN 逾期天數 BETWEEN 61 AND 90 THEN CAST(結餘 AS INT) END,'') '90日～61日', ");
            sb.Append(" ISNULL(CASE WHEN 逾期天數 BETWEEN 91 AND 180 THEN CAST(結餘 AS INT) END,'') '180日～91日', ");
            sb.Append(" ISNULL(CASE WHEN 逾期天數 >180 THEN CAST(結餘 AS INT) END,'') '181日以上' ");
            sb.Append(" FROM (Select       ''''+A.CustID 客戶編號,         ");
            sb.Append(" B.ShortName  客戶名稱           ");
            sb.Append(" , CASE PREPAYDAY WHEN 0 THEN 0 ELSE datediff(day,CAST(CAST(PREPAYDAY AS VARCHAR) AS datetime), @BILLDATE2)  END 逾期天數, ");
            sb.Append(" CAST((case A.Flag when 500 then a.Total+A.Tax  when 595 then a.Total+A.Tax  when 600 then  -(A.Total+A.Tax) when 698 then  -(A.Total+A.Tax) end-A.CashPay-A.VisaPay-A.OtherPay- A.OffSet) AS DECIMAL(12,4)) as 結餘 ");
            sb.Append(" ,CASE WHEN A.SalesMan   IN ('E038','E048','E049') THEN  'PV' ELSE 'TFT' END DTYPE ");
            sb.Append(" From comBillAccounts A             ");
            sb.Append(" Left Join  comCustomer B On A.CustID=B.ID And A.CustFlag=B.Flag        ");
            sb.Append(" Left Join comPerson E On E.PersonID=A.SalesMan           ");
            sb.Append(" Where     A.Flag <> 298 And A.HasCheck = 1 And A.CustFlag =1  ");
            sb.Append(" And (A.Status IN (1, 3))      ");
            sb.Append(" and ((Case When A.Flag IN(297,697,200,600,210,201,698) Then -(A.Total+A.Tax -A.CashPay-A.VisaPay-A.OtherPay)  Else (A.Total+A.Tax -A.CashPay-A.VisaPay-A.OtherPay)      ");
            sb.Append(" End - (A.Offset+A.NoCheckOffSet + A.Discount + A.NoCheckDisCount))<>0)  And A.YearCompressType <> 1     ");
            sb.Append(" AND  A.BILLDATE <= @BILLDATE ) AS A WHERE DTYPE=@DTYPE ) AS A");
            sb.Append(" GROUP BY 客戶編號,客戶名稱");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            string DAY1 = textBox1.Text;
            command.Parameters.Add(new SqlParameter("@BILLDATE", DAY1));
            command.Parameters.Add(new SqlParameter("@DTYPE", DTYPE));
            string DATE = GetMenu.DayS(DAY1);
            command.Parameters.Add(new SqlParameter("@BILLDATE2", GetMenu.DayS(DAY1)));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rdr1");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rdr1"];
        }

        private void ADGETPAY_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.Day();
        }
    }
}

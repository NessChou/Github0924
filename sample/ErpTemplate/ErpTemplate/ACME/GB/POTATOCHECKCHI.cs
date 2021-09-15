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
namespace ACME
{
    public partial class POTATOCHECKCHI : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public POTATOCHECKCHI()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "" && textBox5.Text == "")
            {
                MessageBox.Show("請選擇訂購公司");
                return;
            }

            System.Data.DataTable DT1 = DT();

            if (DT1.Rows.Count > 0)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\GW\\對帳單.xls";

                string ExcelTemplate = FileName;
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                ExcelReport.ExcelReportOutput(DT1, ExcelTemplate, OutPutFile, "N");
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        private System.Data.DataTable DT()
        {


            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT U.FullName 客戶名稱, A.LinkMan 聯絡人,''''+LinkTelephone 電話,U.taxno 統一編號,      ");
            sb.Append(" CASE ISNULL(G.PreInDate,'') WHEN '' THEN '' ELSE SUBSTRING(CAST(G.PreInDate AS VARCHAR),1,4)+'/'+SUBSTRING(CAST(G.PreInDate AS VARCHAR),5,2)+'/'+SUBSTRING(CAST(G.PreInDate AS VARCHAR),7,2) END 到貨日期 ");
            sb.Append(" ,A.CUSTBILLNO 採購單號,  G.ProdID 料號,G.ProdName 品名規格,G.QUANTITY 數量,B.Unit 單位,G.Price 單價,O.TaxAmt 稅,CAST(O.TaxAmt+O.Amount  AS INT) 小計 ");
            sb.Append(" , REPLACE(A.LinkMan,'棉花田-','') 門市,A.BillNO 備註,I.InvoiceNO 發票號碼,I.InvoiceDate 發票日期,TITLE=@TITLE  FROM  OrdBillMain A  Inner Join OrdBillSub G   ");
            sb.Append(" On G.Flag=A.Flag  And G.BillNO=A.BillNO   ");
            sb.Append(" left join ComProdRec O On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO  AND O.Flag =500   ");
            sb.Append(" left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)  ");
            sb.Append(" Left Join comInvoice I On O.BillNO=I.SrcBillNO AND I.Flag =2 ");
            sb.Append(" left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1  ");
            sb.Append(" Left Join comProduct B On B.ProdID=G.ProdID   "); ;
            sb.Append("              WHERE A.Flag =2 ");
            if (textBox5.Text == "")
            {
          
                sb.Append("                    AND U.FullName=@COMPANY ");

                if (textBox6.Text == "" && textBox7.Text == "")
                {
                    sb.Append("              AND A.UserDef1 BETWEEN @CreateDate AND @CreateDate2 ");
                    sb.Append("              AND O.ProdName NOT LIKE '%運費%'  ");
                }
                else
                {
                    sb.Append("              AND  A.BillNO BETWEEN @BILL1 AND @BILL2 ");

                }
                
            }
            else
            {
                sb.Append("              AND A.CUSTBILLNO = @CUSTBILLNO ");
            }


            
            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CreateDate", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CreateDate2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@COMPANY", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@CUSTBILLNO", textBox5.Text));
            command.Parameters.Add(new SqlParameter("@BILL1", textBox6.Text));
            command.Parameters.Add(new SqlParameter("@BILL2", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@TITLE", textBox8.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable DT2()
        {

            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT S.UDef2 運輸單號,CASE B.ConverRate WHEN 0 THEN 0  ELSE CAST(G.QUANTITY/CAST(B.ConverRate AS DECIMAL(6,2)) AS DECIMAL(6,2)) END 箱 ");
            sb.Append(" ,'訂單號碼: '+CAST(A.BillNO AS VARCHAR) 訂單號碼, '收貨人: '+A.LinkMan 收貨人,'收貨地址: '+A.CustAddress 收貨地址");
            sb.Append(" ,'聯繫電話: '+LinkTelephone 聯繫電話, CASE ISNULL(S.UDef1,'') WHEN '' THEN '' ELSE  SUBSTRING(S.UDef1,1,4)+'/'+SUBSTRING(S.UDef1,5,2)+'/'+SUBSTRING(S.UDef1,7,2) END ");
            sb.Append("                          +'  '+ISNULL(A.UserDef2,'')  到貨日期,'訂單日期: '+SUBSTRING(CAST(A.BillDate AS VARCHAR),1,4)+'/'+SUBSTRING(CAST(A.BillDate AS VARCHAR),5,2)+'/'+SUBSTRING(CAST(A.BillDate AS VARCHAR),7,2) 訂單日期");
            sb.Append("                          ,G.ProdName 品名,RANK() OVER (ORDER BY G.ProdName DESC) AS [NO]  FROM  OrdBillMain A  Inner Join OrdBillSub G ");
            sb.Append(" On G.Flag=A.Flag  And G.BillNO=A.BillNO ");
            sb.Append(" left join ComProdRec O On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO  AND O.Flag =500 ");
            sb.Append("     left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500)");
            sb.Append("      left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1");
            sb.Append("      Left Join comProduct B On B.ProdID=G.ProdID ");
            sb.Append("              WHERE A.Flag =2 ");
            sb.Append("              AND A.BillNO=@BillNO ");
            sb.Append("              AND O.ProdName NOT LIKE '%運費%'  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", textBox4.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void POTATOCHECK_Load(object sender, EventArgs e)
        {
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
           
            textBox8.Text = DateTime.Now.ToString("yyyy").ToString()+ "年"+ DateTime.Now.ToString("MM").ToString()+ "月份 交易對帳請款單";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHIPCARD();

            if (LookupValues != null)
            {
                textBox3.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "" && textBox5.Text == "")
            {
                MessageBox.Show("請輸入訂購公司");
                return;
            }

            System.Data.DataTable DT1 = DT();
            if (DT1.Rows.Count > 0)
            {
                dataGridView1.DataSource = DT1;
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox4.Text == "")
            {
                MessageBox.Show("請輸入訂單號碼");
                return;
            }

            System.Data.DataTable DT1 = DT2();

            if (DT1.Rows.Count > 0)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\GW\\出貨簽收單.xls";

                string ExcelTemplate = FileName;
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                ExcelReport.ExcelReportOutput(DT1, ExcelTemplate, OutPutFile, "N");
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox3.Text == "" && textBox5.Text == "")
            {
                MessageBox.Show("請選擇訂購公司");
                return;
            }

            System.Data.DataTable DT1 = DT();

            if (DT1.Rows.Count > 0)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\GW\\對帳單棉花田.xls";

                string ExcelTemplate = FileName;
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                ExcelReport.ExcelReportOutput(DT1, ExcelTemplate, OutPutFile, "N");
            }
            else
            {
                MessageBox.Show("沒有資料");
            }
        }

    }
}

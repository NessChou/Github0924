using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class GBCUST : Form
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GBCUST()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            dataGridView1.DataSource = GetCHO();

        }

        public System.Data.DataTable GetCHO()
        {

            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U.ID 客戶代碼,U.ShortName 客戶簡稱,L.ClassName 類別名稱 ,U.TaxNo 統一編號,U.Telephone1 聯絡電話,[ADDRESS] 發票地址,");
            sb.Append(" CASE U2.RecvWay WHEN 0 THEN '貨到' WHEN 1 THEN '次月' WHEN 2 THEN '月結' WHEN 3 THEN '其他' END 收款方式,U2.DistDays 天  FROM comCustomer U ");
            sb.Append(" Left join comCustTrade U2 On  U.ID=U2.ID AND U2.Flag =1   ");
            sb.Append(" Left Join comCustClass L On L.ClassID =U.ClassID and L.Flag =1   ");
            sb.Append(" Left Join comCustDesc M On U.ID =M.ID and M.Flag =1      ");
            sb.Append(" Left join comCustAddress AD ON (M.AddrID =AD.AddrID AND M.ID=AD.ID )    ");
            //sb.Append(" WHERE U.Flag=1 AND U.ID > '90143-03'  AND SUBSTRING(U.ID,1,2) <>'TW' ORDER BY  U.ID ");
            sb.Append(" WHERE U.Flag=1  AND SUBSTRING(U.ID,1,2) <>'TW' ORDER BY  U.ID ");
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

        private void GBCUST_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetCHO();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }
    }
}

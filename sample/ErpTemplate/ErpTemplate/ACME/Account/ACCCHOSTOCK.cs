using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using CarlosAg.ExcelXmlWriter;
using System.IO;

namespace ACME
{
    public partial class ACCCHOSTOCK : Form
    {
        public ACCCHOSTOCK()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }


        private System.Data.DataTable GetAccountACME(string RefDate1, string RefDate2, string CONN)
        {
            //合計 AS 銷售金額
            SqlConnection connection = new SqlConnection(CONN);

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT T0.PRODID 產品編號,MAX(T0.ProdName) 品名規格,T2.WareHouseName 倉庫,SUM(Quantity) 數量,SUM(Quantity*T0.CAvgCost) 成本  FROM comProduct T0");
            sb.Append(" LEFT JOIN comWareAmount T1 ON ( T0.PRODID=T1.PRODID)");
            sb.Append(" LEFT JOIN comWareHouse T2 ON (T1.WareID=T2.WareHouseID)WHERE Quantity >0");
            sb.Append(" GROUP BY T2.WareHouseName,T0.PRODID");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OJDT");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables[0]; ;


        }
    }
}

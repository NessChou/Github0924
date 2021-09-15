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
    public partial class ScheSap : Form
    {
        private System.Data.DataTable dtCost;
        public ScheSap()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "銷售訂單")
            {
                ViewBatchPayment1();
            }
            else if (comboBox1.Text == "發貨")
            {
                ViewBatchPayment2();
            }
            else if (comboBox1.Text == "借出")
            {
                ViewBatchPayment3();
            }
            else if (comboBox1.Text == "調撥")
            {
                ViewBatchPayment4();
            }
        }

       

        private void ScheSap_Load(object sender, EventArgs e)
        {
             dtCost = MakeTableCombine();
             textBox1.Text = GetMenu.DFirst();
             textBox2.Text = GetMenu.DLast();
             comboBox1.Text = "銷售訂單";
             comboBox2.Text = "未結";
             comboBox3.Text = "請選擇";
             comboBox4.Text = "請選擇";
             comboBox5.Text = "請選擇";
        }


        private void ViewBatchPayment1()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT '銷售訂單' 總類,T0.docentry 單號,T0.linenum 列號,t1.shippingcode SHIPNO,(T4.[SlpName]) 業務人員,(T5.[lastName]+T5.[firstName]) 業助,case cardcode when '0511-00' then 'Choice-'+t3.u_beneficiary when '0257-00' then 'TOP-'+t3.u_beneficiary ELSE t3.cardname END 客戶,");
            sb.Append(" T0.itemcode 產品編號,t0.dscription 品名規格,t14.whsname 倉庫,T0.u_acme_workday+'('+CAST(T2.day AS VARCHAR)+')' 工作天數,");
            sb.Append(" Convert(varchar(8),T0.u_acme_work,112)  排程日期,Convert(varchar(8),T0.u_acme_shipday,112)  離倉日期,");
            sb.Append(" cast(T0.quantity as int) 數量,T1.ETC ETC,T1.ETD ETD,T1.ETA ETA,");
            sb.Append(" T1.receivePlace SHIP_FROM,T1.goalPlace SHIP_TO,T1.tradeCondition TERM,T1.receiveDay 運送方式,");
            sb.Append(" T1.notifyMemo 備註,T0.u_acme_DSCRIPTION SA備註 FROM acmesql02.dbo.rdr1 T0 ");
            sb.Append(" left join (select t0.docentry docentry,t0.linenum linenum,t1.closeDay ETC");
            sb.Append(" ,t1.forecastDay ETD,arriveDay ETA,notifyMemo notifyMemo,buCardcode,receiveDay,boardCountNo,receivePlace,goalPlace,tradeCondition,t1.shippingcode from shipping_main t1");
            sb.Append(" left join shipping_item t0 on (t1.shippingcode=t0.shippingcode) where t0.itemremark='銷售訂單') t1 on (t0.docentry=t1.docentry and t0.linenum=t1.linenum)");
     
            if (comboBox3.Text != "請選擇")
            {
                sb.Append("           left join (select T0.docentry,T0.LINENUM,buCardcode from wh_item4 t0 ");
                sb.Append(" left join WH_main t1 on (t0.shippingcode=t1.shippingcode) where t0.itemremark='銷售訂單') t11 on (t0.docentry=t11.docentry and t0.linenum=t11.linenum) ");
              
            }
            if (comboBox4.Text != "請選擇")
            {
                sb.Append("           left join (select T0.docentry,T0.LINENUM from wh_item t0 ");
                sb.Append("            where t0.itemremark='銷售訂單') t12 on (t0.docentry=t12.docentry and t0.linenum=t12.linenum) ");

            }
            sb.Append("                   left join  acmesqlsp.dbo.WorkDay T2 on (T2.workday=T0.u_acme_workday ) ");
            sb.Append("                   left join  acmesql02.dbo.ORDR T3 on (T0.DOCENTRY=T3.DOCENTRY ) ");
            sb.Append("                   LEFT JOIN ACMESQL02.DBO.OSLP T4 ON T3.SlpCode = T4.SlpCode");
            sb.Append("                   LEFT JOIN ACMESQL02.DBO.OHEM T5 ON T3.OwnerCode = T5.empID");
            sb.Append(" iNNER JOIN ACMESQL02.DBO.OWHS T14 ON T14.whsCode = T0.whscode");
            sb.Append("                   where 1=1 and t3.canceled <> 'Y' AND T3.doctype='I'  ");
            if (comboBox2.Text == "已結")
            {
                sb.Append("and t0.linestatus = 'C'  ");
            }
            else if (comboBox2.Text == "未結")
            {
                sb.Append("and t0.linestatus = 'O' ");
            }
            if (comboBox5.Text == "已結")
            {
                sb.Append("and t1.buCardcode = 'Checked'  ");
            }
            else if (comboBox5.Text == "未結")
            {
                sb.Append("and t1.buCardcode = 'Unchecked' ");
            }
            if (textBox1.Text != "")
            {
                sb.Append("and  Convert(varchar(8),T0.u_acme_work,112) >= '" + textBox1.Text.ToString() + "'  ");
            }
            if (textBox2.Text != "")
            {
                sb.Append("and   Convert(varchar(8),T0.u_acme_work,112)  <=  '" + textBox2.Text.ToString() + "'  ");
            }

            if (textBox3.Text != "")
            {
                sb.Append("and  T0.DOCENTRY  =  '" + textBox3.Text.ToString() + "'  ");
            }
          
             if (comboBox3.Text == "倉庫未結")
            {
                sb.Append("and T0.u_acme_workday = '內銷' and t11.buCardcode = 'Unchecked' ");
            }
            else if (comboBox3.Text == "倉庫已結")
            {
                sb.Append("and T0.u_acme_workday = '內銷' and t11.buCardcode = 'Checked' ");
            }

            if (comboBox4.Text == "倉庫未備")
            {
                sb.Append("and isnull(t12.docentry,'') = '' ");
            }
            else if (comboBox4.Text == "倉庫已備")
            {
                sb.Append("and isnull(t12.docentry,'') <> '' ");
            }
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


            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPayment2()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT '發貨' 總類,T0.docentry 單號,T0.linenum 列號,t1.shippingcode SHIPNO,'' 業務人員, t3.ref2 業助,T3.U_ACME_CUS 客戶,");
            sb.Append(" T0.itemcode 產品編號,t0.dscription 品名規格,t14.whsname 倉庫,'' 工作天數,");
            sb.Append(" Convert(varchar(8),T3.docdate,112)   排程日期,''  離倉日期,");
            sb.Append(" cast(T0.quantity as int) 數量,T1.ETC ETC,T1.ETD ETD,T1.ETA ETA,");
            sb.Append(" T1.receivePlace SHIP_FROM,T1.goalPlace SHIP_TO,T1.tradeCondition TERM,T1.receiveDay 運送方式,");
            sb.Append(" T1.notifyMemo 備註,'' SA備註");
            sb.Append(" FROM acmesql02.dbo.ige1 T0 ");
            sb.Append("          left join (select t0.shippingcode,t0.docentry docentry,t0.linenum linenum,t1.closeDay ETC");
            sb.Append("          ,t1.forecastDay ETD,arriveDay ETA,notifyMemo ,receiveDay,tradeCondition,buCardcode,receivePlace,goalPlace from shipping_item t0 ");
            sb.Append("          left join shipping_main t1 on (t0.shippingcode=t1.shippingcode) where t0.itemremark='發貨單') t1 on (t0.docentry=t1.docentry and t0.linenum=t1.linenum )");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.owhs T7 ON (T0.whscode=T7.whscode)");
            sb.Append(" left join  acmesql02.dbo.Oige T3 on (T0.DOCENTRY=T3.DOCENTRY ) ");
            sb.Append(" iNNER JOIN ACMESQL02.DBO.OWHS T14 ON T14.whsCode = T0.whscode");
            sb.Append(" where t0.u_acme_kind='62880501 - 樣品費' ");
            if (comboBox2.Text == "已結")
            {
                sb.Append("and t1.buCardcode = 'Checked'  ");
            }
            else if (comboBox2.Text == "未結")
            {
                sb.Append("and isnull(T1.buCardcode,'') <> 'Checked' ");
            }
         
            if (textBox1.Text != "")
            {
                sb.Append("and  Convert(varchar(8),T3.docdate,112)  >= '" + textBox1.Text.ToString() + "'  ");
            }
            if (textBox2.Text != "")
            {
                sb.Append("and  Convert(varchar(8),T3.docdate,112)    <=  '" + textBox2.Text.ToString() + "'  ");
            }

            if (textBox3.Text != "")
            {
                sb.Append("and  T0.DOCENTRY  =  '" + textBox3.Text.ToString() + "'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱

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


            dataGridView1.DataSource = ds.Tables[0];

        }

        private void ViewBatchPayment3()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT '借出' 總類,T0.docentry 單號,T0.linenum 列號,t11.shippingcode SHIPNO,t4.slpname 業務人員, '' 業助,t3.cardname 客戶,");
            sb.Append("              T0.itemcode 產品編號,t0.dscription 品名規格,t14.whsname 倉庫,'' 工作天數,");
            sb.Append("              ''  排程日期,''  離倉日期,");
            sb.Append("              cast(T0.quantity as int) 數量,T12.closeDay ETC,T12.forecastDay ETD,T12.arriveDay ETA,");
            sb.Append("              T12.receivePlace SHIP_FROM,T12.goalPlace SHIP_TO,T12.tradeCondition TERM,T12.receiveDay 運送方式,");
            sb.Append("              T12.notifyMemo 備註,'' SA備註");
            sb.Append("              FROM acmesql02.dbo.wtr1 T0 ");
            sb.Append("              left join  dbo.shipping_item T11 on (t0.docentry=T11.docentry and t0.linenum=T11.linenum and t11.itemremark='調撥單') ");
            sb.Append("              left join  dbo.shipping_main T12 on (t11.shippingcode=T12.shippingcode )              ");
            sb.Append("              left join  acmesql02.dbo.Owtr T3 on (T0.DOCENTRY=T3.DOCENTRY ) ");
            sb.Append("                 LEFT JOIN ACMESQL02.DBO.OSLP T4 ON T3.SlpCode = T4.SlpCode");
            sb.Append("              LEFT JOIN ACMESQL02.DBO.owhs T7 ON T0.whscode=T7.whscode");
            sb.Append("              iNNER JOIN ACMESQL02.DBO.OWHS T14 ON T14.whsCode = T0.whscode");
            sb.Append("              where 1=1 and t3.u_acme_kind='1' ");
            if (comboBox2.Text == "已結")
            {
                sb.Append("and t12.buCardcode = 'Checked'  ");
            }
            else if (comboBox2.Text == "未結")
            {
                sb.Append("and isnull(T12.buCardcode,'') <> 'Checked' ");
            }
            if (textBox1.Text != "")
            {
                sb.Append("and  Convert(varchar(8),T3.docdate,112)>= '" + textBox1.Text.ToString() + "'  ");
            }
            if (textBox2.Text != "")
            {
                sb.Append("and   Convert(varchar(8),T3.docdate,112)    <=  '" + textBox2.Text.ToString() + "'  ");
            }
            if (textBox3.Text != "")
            {
                sb.Append("and  T0.DOCENTRY  =  '" + textBox3.Text.ToString() + "'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱

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


            dataGridView1.DataSource = ds.Tables[0];

        }
        private void ViewBatchPayment4()
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("         SELECT '調撥' 總類,T0.docentry 單號,T0.linenum 列號,t11.shippingcode SHIPNO,t4.slpname 業務人員, '' 業助,t3.cardname 客戶,");
            sb.Append("              T0.itemcode 產品編號,t0.dscription 品名規格,t14.whsname 倉庫,'' 工作天數,");
            sb.Append("              ''  排程日期,''  離倉日期,");
            sb.Append("              cast(T0.quantity as int) 數量,T12.closeDay ETC,T12.forecastDay ETD,T12.arriveDay ETA,");
            sb.Append("              T12.receivePlace SHIP_FROM,T12.goalPlace SHIP_TO,T12.tradeCondition TERM,T12.receiveDay 運送方式,");
            sb.Append("              T12.notifyMemo 備註,'' SA備註");
            sb.Append("              FROM acmesql02.dbo.wtr1 T0 ");
            sb.Append("              left join  dbo.shipping_item T11 on (t0.docentry=T11.docentry and t0.linenum=T11.linenum and t11.itemremark='調撥單') ");
            sb.Append("              left join  dbo.shipping_main T12 on (t11.shippingcode=T12.shippingcode )              ");
            sb.Append("              left join  acmesql02.dbo.Owtr T3 on (T0.DOCENTRY=T3.DOCENTRY ) ");
            sb.Append("                 LEFT JOIN ACMESQL02.DBO.OSLP T4 ON T3.SlpCode = T4.SlpCode");
            sb.Append("              LEFT JOIN ACMESQL02.DBO.owhs T7 ON T0.whscode=T7.whscode");
            sb.Append("              iNNER JOIN ACMESQL02.DBO.OWHS T14 ON T14.whsCode = T0.whscode");
            sb.Append("              where 1=1 and t3.u_acme_kind='3' ");
            if (comboBox2.Text == "已結")
            {
                sb.Append("and t12.buCardcode = 'Checked'  ");
            }
            else if (comboBox2.Text == "未結")
            {
                sb.Append("and isnull(T12.buCardcode,'') <> 'Checked' ");
            }
            if (textBox1.Text != "")
            {
                sb.Append("and  Convert(varchar(8),T3.docdate,112)>= '" + textBox1.Text.ToString() + "'  ");
            }
            if (textBox2.Text != "")
            {
                sb.Append("and   Convert(varchar(8),T3.docdate,112)    <=  '" + textBox2.Text.ToString() + "'  ");
            }
            if (textBox3.Text != "")
            {
                sb.Append("and  T0.DOCENTRY  =  '" + textBox3.Text.ToString() + "'  ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱

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


            dataGridView1.DataSource = ds.Tables[0];

        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataRow dr ;
            for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
            {
                string a0 = dataGridView1.SelectedRows[i].Cells[0].Value.ToString();
                string a1 = dataGridView1.SelectedRows[i].Cells[1].Value.ToString();
                string a2 = dataGridView1.SelectedRows[i].Cells[2].Value.ToString();
                string a3 = dataGridView1.SelectedRows[i].Cells[3].Value.ToString();
                string a4 = dataGridView1.SelectedRows[i].Cells[4].Value.ToString();
                string a5 = dataGridView1.SelectedRows[i].Cells[5].Value.ToString();
                string a6 = dataGridView1.SelectedRows[i].Cells[6].Value.ToString();
                string a7 = dataGridView1.SelectedRows[i].Cells[7].Value.ToString();
                string a8 = dataGridView1.SelectedRows[i].Cells[8].Value.ToString();
                string a9 = dataGridView1.SelectedRows[i].Cells[9].Value.ToString();
                string a10 = dataGridView1.SelectedRows[i].Cells[10].Value.ToString();
                string a11 = dataGridView1.SelectedRows[i].Cells[11].Value.ToString();
                string a12 = dataGridView1.SelectedRows[i].Cells[12].Value.ToString();
                string a13 = dataGridView1.SelectedRows[i].Cells[13].Value.ToString();
                string a14 = dataGridView1.SelectedRows[i].Cells[14].Value.ToString();
                string a15 = dataGridView1.SelectedRows[i].Cells[15].Value.ToString();
                string a16 = dataGridView1.SelectedRows[i].Cells[16].Value.ToString();
                string a17 = dataGridView1.SelectedRows[i].Cells[17].Value.ToString();
                string a18 = dataGridView1.SelectedRows[i].Cells[18].Value.ToString();
                string a19 = dataGridView1.SelectedRows[i].Cells[19].Value.ToString();
                string a20 = dataGridView1.SelectedRows[i].Cells[20].Value.ToString();
                string a21 = dataGridView1.SelectedRows[i].Cells[21].Value.ToString();
                string a22 = dataGridView1.SelectedRows[i].Cells[22].Value.ToString();
                dr = dtCost.NewRow();
                dr["總類"] = a0;
                dr["單號"] = a1;
                dr["業務人員"] = a4;
                dr["業助"] = a5;
                dr["客戶"] = a6;
                dr["產品編號"] = a7;
                dr["品名規格"] = a8;
                dr["倉庫"] = a9;
                dr["工作天數"] = a10;
                dr["排程日期"] = a11;
                dr["離倉日期"] = a12;
                dr["數量"] = a13;
                dr["SHIPNO"] = a3;
                dr["ETC"] = a14;
                dr["ETD"] = a15;
                dr["ETA"] = a16;
                dr["SHIP_FROM"] = a17;
                dr["SHIP_TO"] = a18;
                dr["TERM"] = a19;
                dr["運送方式"] = a20;
                dr["備註"] = a21;
                dr["SA備註"] = a22;
                dtCost.Rows.Add(dr);
            }
            dataGridView2.DataSource = dtCost;

        }


        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("總類", typeof(string));
            dt.Columns.Add("單號", typeof(string));
            dt.Columns.Add("業務人員", typeof(string));
            dt.Columns.Add("業助", typeof(string));
            dt.Columns.Add("客戶", typeof(string));
            dt.Columns.Add("產品編號", typeof(string));
            dt.Columns.Add("品名規格", typeof(string));
            dt.Columns.Add("倉庫", typeof(string));
            dt.Columns.Add("工作天數", typeof(string));
            dt.Columns.Add("排程日期", typeof(string));
            dt.Columns.Add("離倉日期", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("SHIPNO", typeof(string));
            dt.Columns.Add("ETC", typeof(string));
            dt.Columns.Add("ETD", typeof(string));
            dt.Columns.Add("ETA", typeof(string));
            dt.Columns.Add("SHIP_FROM", typeof(string));
            dt.Columns.Add("SHIP_TO", typeof(string));
            dt.Columns.Add("TERM", typeof(string));
            dt.Columns.Add("運送方式", typeof(string));
            dt.Columns.Add("備註", typeof(string));
            dt.Columns.Add("SA備註", typeof(string));
            return dt;
        }

 

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);

        }


     
    

       
    

       
      

      
      
       
    }
}
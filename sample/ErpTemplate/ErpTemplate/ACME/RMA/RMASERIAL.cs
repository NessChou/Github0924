using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using CarlosAg.ExcelXmlWriter;
namespace ACME
{

    public partial class RMASERIAL : Form
    {
        int S1 = 0;
        string FileName = "";

        public RMASERIAL()
        {
            InitializeComponent();
        }




        public void AddINVOUT(string SHIPPING, string PART, string INVOICE, string CARTON, string USERID, string CELLCHIP, string RMANO, string VRMANO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_INVOICEOUT(SHIPPING,PART,INVOICE,CARTON,INSERTDATE,USERID,CELLCHIP,RMANO,VRMANO) values(@SHIPPING,@PART,@INVOICE,@CARTON,@INSERTDATE,@USERID,@CELLCHIP,@RMANO,@VRMANO)", connection);
            command.CommandType = CommandType.Text;
           
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));
            command.Parameters.Add(new SqlParameter("@PART", PART));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@INSERTDATE", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@USERID", USERID));
            command.Parameters.Add(new SqlParameter("@CELLCHIP", CELLCHIP));
            command.Parameters.Add(new SqlParameter("@RMANO", RMANO));
            command.Parameters.Add(new SqlParameter("@VRMANO", VRMANO));
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void AddINVOUT2(string SHIPPING)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_INVOICEOUT2(SHIPPING) values(@SHIPPING)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void TRUNCATE()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE RMA_INVOICEOUT2", connection);
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
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }


        public void DELINVOUT(string INVOICE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE RMA_INVOICEOUT WHERE INVOICE=@INVOICE ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));


            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }

        private System.Data.DataTable GetOrderData20AP(string TYPE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


  
            if (TYPE == "2")
            {
                sb.Append(" SELECT PART SHIPPING_PART_NO,SHIPPING  SHIPPING_NO ,INVOICE INVOICE_NO,CARTON CARTON_NO,CELLCHIP 'CELL_CHIP_ID' FROM RMA_INVOICEOUT where (SUBSTRING(INVOICE,1,3) = ('BVT')  OR SUBSTRING(INVOICE,1,2) = ('FH')) ");
            }
            if (TYPE == "3")
            {
                sb.Append(" SELECT PART SHIPPING_PART_NO,SHIPPING  SHIPPING_NO ,INVOICE INVOICE_NO,CARTON CARTON_NO,CELLCHIP 'CELL_CHIP_ID' FROM RMA_INVOICEOUT where SUBSTRING(INVOICE,1,2) IN ('TL','LB') ");
            }


            if (textBox1.Text != "")
            {
                sb.Append("  and SHIPPING  = '" + textBox1.Text.ToString() + "' ");
            }
            if (textBox2.Text != "")
            {
                sb.Append("  and (INVOICE  = '" + textBox2.Text.ToString() + "' )  ");
            }
            if (textBox3.Text != "")
            {
                sb.Append("  and CARTON  = '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox4.Text != "")
            {
                sb.Append("  and CELLCHIP  = '" + textBox4.Text.ToString() + "' ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
     


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

            return ds.Tables[0];

        }


        private System.Data.DataTable GetOrderDataD2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PART SHIPPING_PART_NO,SHIPPING SHIPPING_NO ,INVOICE INVOICE_NO,CARTON CARTON_NO,CELLCHIP,RMANO,VRMANO  FROM RMA_INVOICEOUT where SUBSTRING(INVOICE,1,3) <> ('BVT') and SUBSTRING(INVOICE,1,2) <> ('FH') and SUBSTRING(INVOICE,1,2) NOT IN ('TL','LB') ");

            if (textBox1.Text != "")
            {
                sb.Append("  and SHIPPING  = '" + textBox1.Text.ToString() + "' ");
            }
            if (textBox2.Text != "")
            {
                sb.Append("  and (INVOICE  = '" + textBox2.Text.ToString() + "' )  ");
            }
            if (textBox3.Text != "")
            {
                sb.Append("  and CARTON  = '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox4.Text != "")
            {
                sb.Append("  and CELLCHIP  = '" + textBox4.Text.ToString() + "' ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData20AP4(string SHIPPING)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();



            sb.Append(" SELECT PART SHIPPING_PART_NO,SHIPPING SHIPPING_NO,INVOICE INVOICE_NO,CARTON CARTON_NO,CELLCHIP 'CELL_CHIP_ID' FROM RMA_INVOICEOUT where SUBSTRING(INVOICE,1,3) <> ('BVT') and SUBSTRING(INVOICE,1,2) <> ('FH') and SUBSTRING(INVOICE,1,2) NOT IN ('TL','LB') AND SHIPPING LIKE  '%" + SHIPPING + "%' ");
            
   
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;

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

            return ds.Tables[0];

        }


        private System.Data.DataTable GetOrderData20AP5(string SHIPPING)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();



            sb.Append(" SELECT PART SHIPPING_PART_NO,SHIPPING SHIPPING_NO,INVOICE INVOICE_NO,CARTON CARTON_NO,CELLCHIP 'CELL_CHIP_ID' FROM RMA_INVOICEOUT where (SUBSTRING(INVOICE,1,3) = ('BVT') OR SUBSTRING(INVOICE,1,2) = ('FH')) AND SHIPPING NOT IN (SELECT SHIPPING FROM RMA_CHIP3) AND SHIPPING =  '" + SHIPPING + "'  ");
        


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



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

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData20AP2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                   SELECT RMANO 'RMA NO',CARDNAME Customer,Model,Ver,U_S_SEQ 'Serial Number',U_U_ACME_JUDGE 'ACME Judge',VENDER 'Vender RMA No'");
            sb.Append("                       FROM acmesqlsp.dbo.rma_MAINF T0 ");
            sb.Append("                    LEFT JOIN acmesqlsp.dbo.RMA_INVOICEF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("                     LEFT JOIN acmesqlsp.dbo.RMA_CTR1 T2 ON (T1.RMANO=T2.U_RMA_NO COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.DOCDATE=T2.ManufSN ) WHERE 1=1  ");
            sb.Append("  and U_S_SEQ  = '" + textBox1.Text.ToString() + "' ");

            sb.Append("         union all ");
            sb.Append("                   SELECT RMANO 'RMA NO',CARDNAME Customer,Model,Ver,U_S_SEQ 'Serial Number',U_U_ACME_JUDGE 'ACME Judge',VENDER 'Vender RMA No'");
            sb.Append("                       FROM acmesqlspdrs.dbo.rma_MAINF T0 ");
            sb.Append("                    LEFT JOIN acmesqlspdrs.dbo.RMA_INVOICEF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("                     LEFT JOIN acmesqlspdrs.dbo.RMA_CTR1 T2 ON (T1.RMANO=T2.U_RMA_NO COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.DOCDATE=T2.ManufSN ) WHERE 1=1  ");
            sb.Append("  and U_S_SEQ  = '" + textBox1.Text.ToString() + "' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
    

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

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData20AP3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                     SELECT RMANO 'RMA NO',DOCDATE 還貨日  FROM acmesqlsp.dbo.rma_MAINF T0 ");
            sb.Append("                     LEFT JOIN acmesqlsp.dbo.RMA_INVOICEF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("                     LEFT JOIN acmesqlsp.dbo.RMA_CTR1 T2 ON (T1.RMANO=T2.U_RMA_NO COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.DOCDATE=T2.ManufSN ) WHERE 1=1  ");
            sb.Append("  and U_S_SEQ  = '" + textBox1.Text.ToString() + "' ");
            sb.Append("         union all ");
            sb.Append("                     SELECT RMANO 'RMA NO',DOCDATE 還貨日  FROM acmesqlspdrs.dbo.rma_MAINF T0 ");
            sb.Append("                     LEFT JOIN acmesqlspdrs.dbo.RMA_INVOICEF T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("                     LEFT JOIN acmesqlspdrs.dbo.RMA_CTR1 T2 ON (T1.RMANO=T2.U_RMA_NO COLLATE Chinese_Taiwan_Stroke_CI_AS AND T0.DOCDATE=T2.ManufSN ) WHERE 1=1  ");
            sb.Append("  and U_S_SEQ  = '" + textBox1.Text.ToString() + "' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



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

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData21AP(string TYPE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            if (TYPE == "1")
            {
                sb.Append(" SELECT PART AUO進貨種類,count(*) Qty FROM RMA_INVOICEOUT where SUBSTRING(INVOICE,1,3) <> ('BVT') and SUBSTRING(INVOICE,1,2) <> ('FH') and SUBSTRING(INVOICE,1,2) NOT IN ('TL','LB') ");

            }
            if (TYPE == "2")
            {
                sb.Append(" SELECT PART 景智產種類,count(*) Qty FROM RMA_INVOICEOUT where (SUBSTRING(INVOICE,1,3) = ('BVT') OR SUBSTRING(INVOICE,1,2) = ('FH')) ");
            }
            if (TYPE == "3")
            {
                sb.Append(" SELECT PART '天樂/連輝產種類',count(*) Qty FROM RMA_INVOICEOUT where  SUBSTRING(INVOICE,1,2) IN ('TL','LB') ");
            }

            if (textBox1.Text != "")
            {
                sb.Append("  and SHIPPING  = '" + textBox1.Text.ToString() + "' ");
            }
            if (textBox2.Text != "")
            {
                sb.Append("  and INVOICE  = '" + textBox2.Text.ToString() + "' ");
            }
            if (textBox3.Text != "")
            {
                sb.Append("  and CARTON  = '" + textBox3.Text.ToString() + "' ");
            }
            if (textBox4.Text != "")
            {
                sb.Append("  and CELLCHIP  = '" + textBox4.Text.ToString() + "' ");
            }
            sb.Append("  group by PART  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@insertdate", GetMenu.Day()));
            command.CommandTimeout = 0;

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

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData21AP1(string SHIPPING)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PART 景智產種類,count(*) Qty FROM RMA_INVOICEOUT where (SUBSTRING(INVOICE,1,3) = ('BVT') OR SUBSTRING(INVOICE,1,2) = ('FH')) AND SHIPPING NOT IN (SELECT SHIPPING FROM RMA_CHIP3) AND SHIPPING =  '" + SHIPPING + "' ");    
            sb.Append("  group by PART  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData21AP12(string SHIPPING)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PART AUO進貨種類,count(*) Qty FROM RMA_INVOICEOUT where SHIPPING LIKE  '%" + SHIPPING + "%'  AND SUBSTRING(INVOICE,1,3) <> ('BVT') and SUBSTRING(INVOICE,1,2) <> ('FH') and SUBSTRING(INVOICE,1,2) NOT IN ('TL','LB')  ");
            sb.Append("  group by PART  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData21AP2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT AUO AUO進貨種類,SUM(Qty) Qty FROM RMA_CHIP   GROUP BY AUO ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


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

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData21AP22()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT AUO 景智產種類,SUM(Qty) Qty FROM RMA_CHIP2  GROUP BY AUO ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


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

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData22AP()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("    select  Convert(varchar(10),(t0.docdate),112) 日期,t0.cardname 廠商名稱");
            sb.Append("           ,t1.ItemCode 產品編號,t1.dscription 品名規格,SUM(cast(t1.quantity as int)) 數量,");
            sb.Append("           LTRIM(RTRIM(T0.U_ACME_INV)) Invoice from opdn t0 ");
            sb.Append("           LEFT JOIN PDN1 T1 ON(T0.docentry=t1.docentry)  INNER  JOIN [dbo].[OITM] TA  ON  TA.[ItemCode] = T1.ItemCode");
            sb.Append("           WHERE   ISNULL(TA.U_GROUP,'') <> 'Z&R-費用類群組' ");


            if (textBox2.Text != "")
            {
                sb.AppendFormat(" and ( t0.u_acme_inv like '%{0}%') ", textBox2.Text.ToString());
            }
       

            sb.Append(" GROUP BY  Convert(varchar(10),(t0.docdate),112) ,t0.cardname ");
            sb.Append("           ,t1.ItemCode ,t1.dscription ,");
            sb.Append("           LTRIM(RTRIM(T0.U_ACME_INV)) ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@insertdate", GetMenu.Day()));


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

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData23AP(string INVOICE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT INVOICE FROM RMA_INVOICEOUT where 1=1 AND INVOICE=@INVOICE ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));


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

            return ds.Tables[0];

        }
        private System.Data.DataTable GetRMA_INVOICEOUT2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T1.PART SHIPPING_PART_NO,T0.SHIPPING SHIPPING_NO,T1.INVOICE INVOICE_NO,T1.CARTON CARTON_NO FROM RMA_INVOICEOUT2 T0 LEFT JOIN  RMA_INVOICEOUT T1 ON (T0.SHIPPING=T1.SHIPPING) ORDER BY  T0.ID");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;



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

            return ds.Tables[0];

        }
  
        public void AddCHIP(string AUO,int QTY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_CHIP(AUO,QTY) values(@AUO,@QTY)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AUO", AUO));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));

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
        public void AddCHIP2(string AUO, int QTY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_CHIP2(AUO,QTY) values(@AUO,@QTY)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AUO", AUO));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));

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
        public void AddCHIP3(string SHIPPING)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_CHIP3(SHIPPING) values(@SHIPPING)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPING", SHIPPING));
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
        public void TRUNCATECHIP()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("truncate table RMA_CHIP truncate table RMA_CHIP2  truncate table RMA_CHIP3 ", connection);
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

        private void button3_Click(object sender, EventArgs e)
        {
            //try
            //{


                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\EXCEL2\\工程匯入\\";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {
                    
                    FileInfo filess = new FileInfo(file);
                    string dd = filess.Name.ToString();
                    int ad = dd.LastIndexOf(".");
                    string PanelName = dd.Substring(0, ad).ToString();


                    GetINVOICE(file, "");

                              
                }


                MessageBox.Show("上傳完成");
            
            //}
            //catch (Exception ex)
            //{
              
            //    MessageBox.Show(ex.Message);
            //}
        }


        private void GetINVOICE(string ExcelFile,string CHIP)
        {
            S1 = 0;

            int N1 = 0;
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();
            string NAME = excelSheet.Name;
            
                   System.Data.DataTable T1 = GetOrderData23AP(NAME);
                   if (T1.Rows.Count > 0)
                   {
                       DialogResult result;
                       result = MessageBox.Show("系統上已有相同資料，請確認是否繼續執行", "YES/NO", MessageBoxButtons.YesNo);
                       if (result == DialogResult.Yes)
                       {
                           DELINVOUT(NAME);
                       }
                       else
                       {
                           return;
                       }
                   }

                   label4.Text = "EXCEL : "+NAME + "處理中";
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id1;
            string id2;
            string id3;
            string id4;
            string id5;
            string id6;
            for (int i = 2; i <= iRowCnt; i++)
            {

                if (S1 == 1)
                {
                    return;
                }

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id1 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                id2 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                id3 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                id4 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                id5 = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                id6 = range.Text.ToString().Trim();

                if (id1 == "" && id2 == "" && id3 == "" && id4 == "")
                {
                    return;
                }
                if (CHIP == "CHIP")
                {
                    AddINVOUT(id2, id1, id3, id4, fmLogin.LoginID.ToString(), id5, "", "");
                }
                else
                {
                    AddINVOUT(id2, id1, id3, id4, fmLogin.LoginID.ToString(), "", id5, id6);
                }
            }

   
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


        private void button5_Click(object sender, EventArgs e)
        {
       
             if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToCSV2(dataGridView2, Environment.CurrentDirectory + @"\" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            }
        
        }


        private void WH(CarlosAg.ExcelXmlWriter.Workbook book, DataGridView DGV, string DD)
        {



           CarlosAg.ExcelXmlWriter.Worksheet sheet = book.Worksheets.Add(DD);
            WorksheetRow headerRow = sheet.Table.Rows.Add();
            for (int i = 0; i < DGV.Columns.Count ; i++)
            {
                headerRow.Cells.Add(DGV.Columns[i].HeaderText, DataType.String, "headerStyleID");
            }

            for (int i = 0; i < DGV.Rows.Count-1; i++)
            {

                DataGridViewRow row = DGV.Rows[i];
                WorksheetRow rowS = sheet.Table.Rows.Add();

                for (int j = 0; j < row.Cells.Count; j++)
                {

                    DataGridViewCell cell = row.Cells[j];

                    //if (j == 0 || j == 1)
                    //{
                        rowS.Cells.Add(cell.Value.ToString(), DataType.String, "workbookStyleID");
                   // }
                    //else
                    //{
                    //    rowS.Cells.Add(cell.Value.ToString(), DataType.Number, "workbookStyleID2");
                    //}
                    rowS.AutoFitHeight = true;
                    rowS.Table.DefaultColumnWidth = 100;

                }

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TRUNCATECHIP();

            if (textBox1.Text == "" && textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == "")
            {
                MessageBox.Show("請輸入查詢條件");
                return;
            }
       //     dataGridView8.DataSource = GetOrderData20AP("2");
            dataGridView11.DataSource = GetOrderData21AP("3");
            dataGridView7.DataSource = GetOrderData20AP("3");
            
            T3();
            T4();

            T1();
            T2(); 

            dataGridView5.DataSource = GetOrderData20AP3();
            dataGridView6.DataSource = GetOrderData20AP2();

 

            if (textBox2.Text != "")
            {
                dataGridView3.DataSource = GetOrderData22AP();
            }
        }
        public void T1()
        {
            System.Data.DataTable dt = GetOrderData21AP("1");

           
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                AddCHIP(dt.Rows[i]["AUO進貨種類"].ToString(), Convert.ToInt16(dt.Rows[i]["Qty"]));
            }


            if (dataGridView8.Rows.Count > 0)
            {

                int i = this.dataGridView8.Rows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    string K1 = dataGridView8.Rows[iRecs].Cells["CELL_CHIP_ID"].Value.ToString();
                    if (!String.IsNullOrEmpty(K1))
                    {
                        System.Data.DataTable G1 = GetOrderData21AP12(K1);
                        if (G1.Rows.Count > 0)
                        {
                            for (int S = 0; S <= G1.Rows.Count - S; S++)
                            {
                                AddCHIP(G1.Rows[S]["AUO進貨種類"].ToString(), Convert.ToInt32(G1.Rows[S]["Qty"]));
                            }
                        }
                    }
                }

            }
            dataGridView1.DataSource = GetOrderData21AP2();
        }

        public void T4()
        {
            System.Data.DataTable dt = GetOrderData21AP("2");


            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                AddCHIP2(dt.Rows[i]["景智產種類"].ToString(), Convert.ToInt16(dt.Rows[i]["Qty"]));
            }


            if (dataGridView7.Rows.Count > 0)
            {

                int i = this.dataGridView7.Rows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    string K1 = dataGridView7.Rows[iRecs].Cells[1].Value.ToString();

                    System.Data.DataTable G1 = GetOrderData21AP1(K1);
                    if (G1.Rows.Count > 0)
                    {
                        for (int S = 0; S <= G1.Rows.Count - S; S++)
                        {
                            AddCHIP2(G1.Rows[S]["景智產種類"].ToString(), Convert.ToInt16(G1.Rows[S]["Qty"]));
                        }
                    }
                }

            }

            dataGridView10.DataSource = GetOrderData21AP22();
        }
        public void T2()
        {
            System.Data.DataTable dt = GetOrderDataD2();
            System.Data.DataTable dtCost = MakeTableCombine2();
            DataRow dr = null;
        
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                dr["SHIPPING_PART_NO"] = dt.Rows[i]["SHIPPING_PART_NO"].ToString();
                dr["SHIPPING_NO"] = dt.Rows[i]["SHIPPING_NO"].ToString();
                dr["INVOICE_NO"] = dt.Rows[i]["INVOICE_NO"].ToString();
                dr["CARTON_NO"] = dt.Rows[i]["CARTON_NO"].ToString();
                dr["CELL_CHIP_ID"] = dt.Rows[i]["CELLCHIP"].ToString();
                dr["RMA_NO"] = dt.Rows[i]["RMANO"].ToString();
                dr["VENDER_RMA_NO"] = dt.Rows[i]["VRMANO"].ToString();
                dtCost.Rows.Add(dr);
            }

            if (dataGridView8.Rows.Count > 0)
            {

                int i = this.dataGridView8.Rows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    string K1 = dataGridView8.Rows[iRecs].Cells["CELL_CHIP_ID"].Value.ToString();
                    if (!String.IsNullOrEmpty(K1))
                    {
                        System.Data.DataTable G1 = GetOrderData20AP4(K1);
                        if (G1.Rows.Count > 0)
                        {
                            for (int S = 0; S <= G1.Rows.Count - S; S++)
                            {
                                dr = dtCost.NewRow();
                                dr["SHIPPING_PART_NO"] = G1.Rows[S]["SHIPPING_PART_NO"].ToString();
                                dr["SHIPPING_NO"] = G1.Rows[S]["SHIPPING_NO"].ToString();
                                dr["INVOICE_NO"] = G1.Rows[S]["INVOICE_NO"].ToString();
                                dr["CARTON_NO"] = G1.Rows[S]["CARTON_NO"].ToString();
                                dr["CELL_CHIP_ID"] = G1.Rows[S]["CELL_CHIP_ID"].ToString();
                                dtCost.Rows.Add(dr);
                            }
                        }
                    }
                }

            }
            dataGridView2.DataSource = dtCost;
        }

        public void T3()
        {
            System.Data.DataTable dt = GetOrderData20AP("2");
            System.Data.DataTable dtCost = MakeTableCombine3();
            DataRow dr = null;

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                dr["SHIPPING_PART_NO"] = dt.Rows[i]["SHIPPING_PART_NO"].ToString();
                dr["SHIPPING_NO"] = dt.Rows[i]["SHIPPING_NO"].ToString();
                dr["INVOICE_NO"] = dt.Rows[i]["INVOICE_NO"].ToString();
                dr["CARTON_NO"] = dt.Rows[i]["CARTON_NO"].ToString();
                dr["CELL_CHIP_ID"] = dt.Rows[i]["CELL_CHIP_ID"].ToString();
                dtCost.Rows.Add(dr);
                AddCHIP3(dt.Rows[i]["SHIPPING_NO"].ToString());
            }

            if (dataGridView7.Rows.Count > 0)
            {

                int i = this.dataGridView7.Rows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    string K1 = dataGridView7.Rows[iRecs].Cells[1].Value.ToString();

                    System.Data.DataTable G1 = GetOrderData20AP5(K1);
                    if (!String.IsNullOrEmpty(K1))
                    {
                        if (G1.Rows.Count > 0)
                        {
                            for (int S = 0; S <= G1.Rows.Count - S; S++)
                            {
                                dr = dtCost.NewRow();
                                dr["SHIPPING_PART_NO"] = G1.Rows[S]["SHIPPING_PART_NO"].ToString();
                                dr["SHIPPING_NO"] = G1.Rows[S]["SHIPPING_NO"].ToString();
                                dr["INVOICE_NO"] = G1.Rows[S]["INVOICE_NO"].ToString();
                                dr["CARTON_NO"] = G1.Rows[S]["CARTON_NO"].ToString();
                                dr["CELL_CHIP_ID"] = G1.Rows[S]["CELL_CHIP_ID"].ToString();
                                dtCost.Rows.Add(dr);
                            }
                        }
                    }
                }

            }
       
            dataGridView8.DataSource = dtCost;
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("AUO進貨種類", typeof(string));
            dt.Columns.Add("Qty", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombine2()
        {
         
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("SHIPPING_PART_NO", typeof(string));
            dt.Columns.Add("SHIPPING_NO", typeof(string));
            dt.Columns.Add("INVOICE_NO", typeof(string));
            dt.Columns.Add("CARTON_NO", typeof(string));
            dt.Columns.Add("CELL_CHIP_ID", typeof(string));
            dt.Columns.Add("RMA_NO", typeof(string));
            dt.Columns.Add("VENDER_RMA_NO", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombine3()
        {

            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("SHIPPING_PART_NO", typeof(string));
            dt.Columns.Add("SHIPPING_NO", typeof(string));
            dt.Columns.Add("INVOICE_NO", typeof(string));
            dt.Columns.Add("CARTON_NO", typeof(string));
            dt.Columns.Add("CELL_CHIP_ID", typeof(string));
            return dt;
        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                ExcelReport.GridViewToExcel(dataGridView3);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                TRUNCATE();
                string F = opdf.FileName;
                GetExcelContentGD44(F);

                dataGridView4.DataSource = GetRMA_INVOICEOUT2();
            }
        }
        private void GetExcelContentGD44(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

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
                id = range.Text.ToString();


                if (!String.IsNullOrEmpty(id))
                {
                    AddINVOUT2(id);

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

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView4);
        }

        private void GetExcelProduct(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();




            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
            excelSheet.Activate();
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {
                string SERIAL_NO;
                string CART_NO;
                string CART_NO2 = "";
   
                //第一行要
                for (int iRecord = 3; iRecord <= iRowCnt; iRecord++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    CART_NO = range.Text.ToString().Trim();
   

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    SERIAL_NO = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(SERIAL_NO))
                    {

                            if (!String.IsNullOrEmpty(SERIAL_NO))
                            {
                                if (String.IsNullOrEmpty(CART_NO))
                                {
                                    CART_NO = CART_NO2;
                                }
                            }
 
                            AddINVCHIP(SERIAL_NO, CART_NO);
                             CART_NO2 = CART_NO;
            
                        }

                }


                Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(3);
                excelSheet2.Activate();
                int iRowCnt2 = excelSheet2.UsedRange.Cells.Rows.Count;
                int iColCnt2 = excelSheet2.UsedRange.Cells.Columns.Count;



                string SERIAL_NO2;
                string CHIPID;

                //第一行要
                for (int iRecord = 2; iRecord <= iRowCnt2; iRecord++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    CHIPID = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    SERIAL_NO2 = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(SERIAL_NO2))
                    {
                        if (!String.IsNullOrEmpty(CHIPID))
                        {


                            AddINVCHIP2(SERIAL_NO2, CHIPID);

                        }

                    }

                }


            }
            finally
            {


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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();

            }


        }




        public void AddINVCHIP(string SHIPNO, string CARTNO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_INVOICECHIP(SHIPNO,CARTNO) values(@SHIPNO,@CARTNO)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPNO", SHIPNO));
            command.Parameters.Add(new SqlParameter("@CARTNO", CARTNO));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }
        public void AddINVCHIP2(string SHIPNO, string CHIPNO)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_INVOICECHIP2(SHIPNO,CHIPNO) values(@SHIPNO,@CHIPNO)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SHIPNO", SHIPNO));
            command.Parameters.Add(new SqlParameter("@CHIPNO", CHIPNO));

            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    S1 = 1;
                    MessageBox.Show(ex.Message);
                }
            }
            finally
            {
                connection.Close();
            }

        }


        private void button12_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                try
                {

                    string FileNameS = openFileDialog1.FileName;

                    string AA = Path.GetDirectoryName(FileNameS) + "\\" +

                 DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileNameS);
                    GetExcelProduct(FileNameS, AA);


                }
                finally
                {
                    Cursor = Cursors.Default;
                }

            }
        }
        private void GetExcelProduct(string ExcelFile, string NewFileName)
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

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Microsoft.Office.Interop.Excel.Range range = null;

            // MessageBox.Show(資產.ToString());

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;



                string ff = "";
                string ss = "";

                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    string AA = range.Text.ToString().Trim();
                    if (AA == "")
                    {
                        ff = ss;
                    }
                    else
                    {
                        ff = AA;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Value2 = ff.ToString();
                    ss = ff;


                    //if (A2 != F1)
                    //{
                    //    int F = YEARS.Length;
                    //    string TYEAR = YEARS.Substring(0, 4);
                    //    string TMON = YEARS.Substring(5, F - 5);
                    //    T1 = GetCHO3ANDNOTCLOSE2(TYEAR, TMON, MANS, TYPE);
                    //    K2 = T1.Rows[0][0].ToString();
                    //}
                    //else
                    //{
                    //    int KH = 0;
                    //    for (int A2S = 5; A2S <= F1 - 1; A2S++)
                    //    {
                    //        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, A2S]);
                    //        range.Select();
                    //        string G = range.Text.ToString().Trim().Replace(",", "");
                    //        if (String.IsNullOrEmpty(G))
                    //        {
                    //            G = "0";
                    //        }
                    //        KH += Convert.ToInt16(G);
                    //    }

                    //    K2 = KH.ToString();
                    //}

                    //if (!String.IsNullOrEmpty(K2))
                    //{
                    //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, A2]);
                    //    range.Value2 = K2;
                    //    if (TYPE == "P")
                    //    {
                    //        KS += Convert.ToInt16(K2);
                    //    }
                    //    if (TYPE == "C")
                    //    {
                    //        K1 += Convert.ToInt16(K2);
                    //    }

                    //}
                    //range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, A2]);
                    //if (AA == "雞 合計")
                    //{
                    //    if (A2 == F1)
                    //    {
                    //        range.Value2 = (K1 / 2).ToString();
                    //    }
                    //    else
                    //    {
                    //        range.Value2 = K1.ToString();
                    //    }
                    //}
                    //if (AA == "豬 合計")
                    //{

                    //    if (A2 == F1)
                    //    {
                    //        range.Value2 = (KS / 2).ToString();
                    //    }
                    //    else
                    //    {
                    //        range.Value2 = KS.ToString();
                    //    }
                    //}
                    //if (AA == "總計")
                    //{
                    //    if (A2 == F1)
                    //    {
                    //        range.Value2 = ((KS + K1) / 3).ToString();
                    //    }
                    //    else
                    //    {
                    //        range.Value2 = (KS + K1).ToString();
                    //    }

                    //}



                    // }



                }

            }
            finally
            {




                try
                {
                    //excelSheet.SaveAs(NewFileName, XlFileFormat., oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    //                excelSheet.SaveAs(NewFileName, XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
                    //Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                    //Excel.XlSaveConflictResolution.xlUserResolution, true,
                    //Missing.Value, Missing.Value, Missing.Value);
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
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);

                System.Diagnostics.Process.Start(NewFileName);
            }

        }

        private void button8_Click(object sender, EventArgs e)
        {

            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string OutPutFile = lsAppDir + "\\EXCEL2\\工程匯入\\";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {

                FileInfo filess = new FileInfo(file);
                string dd = filess.Name.ToString();
                int ad = dd.LastIndexOf(".");
                string PanelName = dd.Substring(0, ad).ToString();


                GetINVOICE(file, "CHIP");


            }


            MessageBox.Show("上傳完成");
        }

 
    }
}
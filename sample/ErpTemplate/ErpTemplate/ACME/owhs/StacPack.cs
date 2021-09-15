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


namespace ACME
{
    public partial class StacPack : Form
    {
        public StacPack()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("叫匡拒郎");
            }
            else
            {
                TRUNCATE();
                string F = opdf.FileName;
                    GetExcelContentGD44(F);
                

            }
        }

        private void GetExcelContentGD44(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

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
            string id2 = "";
            string id3 = "";
            string id4 = "";
            string id5 = "";

            int u = 0;
            int v = 0;
   
     
            for (int i = 7; i <= iRowCnt; i++)
            {




                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id = range.Text.ToString();
   

                if (!String.IsNullOrEmpty(id))
                {
                    DateTime dd = Convert.ToDateTime(id);

                    string df = dd.ToString("yyyyMMdd");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                    id2 = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                    id3 = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 10]);
                    id4 = range.Text.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 11]);
                    id5 = range.Text.ToString();
    

                    string a1=DateTime.Now.ToString("yyyyMMdd");
                    AddAUOGD4(a1, df, id2, id3, id4, id5);
             
                    


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
            MessageBox.Show("蹲Θ");
        }

        public void AddAUOGD4(string INSDATE,string DOCDATE,string INVOICE,string ITEMCODE,string PACK,string QTY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_PACK(INSDATE,DOCDATE,INVOICE,ITEMCODE,PACK,QTY) values(@INSDATE,@DOCDATE,@INVOICE,@ITEMCODE,@PACK,@QTY)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@INSDATE", INSDATE));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@INVOICE", INVOICE));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@PACK", PACK));
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

        public void AddWH(string AA,string BB)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE  OITM SET U_GROUP=@AA WHERE ITEMCODE=@BB", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@AA", AA));
            command.Parameters.Add(new SqlParameter("@BB", BB));



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
        public void TRUNCATE()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE WH_PACK", connection);
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
        private System.Data.DataTable Get3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("          SELECT * FROM (             SELECT CASE WHEN SUBSTRING(ITEMCODE,1,1) LIKE '[A-Z]%' AND ");
            sb.Append("                  SUBSTRING(ITEMCODE,2,1) LIKE '[0-9]%' AND ");
            sb.Append("                  SUBSTRING(ITEMCODE,3,1) LIKE '[0-9]%'");
            sb.Append("                 AND SUBSTRING(ITEMCODE,4,1) LIKE '[0-9]%' THEN  Substring ([ItemCode],1,9)  ELSE ");
            sb.Append("          Substring ([ItemCode],2,8) END  MODEL,Substring([ItemCode],12,1) VER,DOCDATE DATE,INVOICE 'AU INV#',ITEMCODE 'ACME P/N',cast(CASE PACK WHEN '' THEN 0 ELSE PACK END as decimal)  '狾计',cast(ISNULL(CASE QTY WHEN '' THEN 0 ELSE QTY END,0) as decimal) 计秖");
            sb.Append("                       ,CASE WHEN (ISNULL(T1.计秖,0)-CASE WHEN CAST((CASE QTY WHEN '' THEN 0 ELSE QTY END) AS DECIMAL)=0 THEN 0");
            sb.Append("                       WHEN CAST(CASE PACK WHEN '' THEN 0 ELSE PACK END AS DECIMAL)=0 THEN 0 ELSE CAST((CASE QTY WHEN '' THEN 0 ELSE QTY END) AS DECIMAL)/CAST(CASE PACK WHEN '' THEN 0 ELSE PACK END AS DECIMAL) END ) <> 0 THEN '钵盽' END '钵盽',");
            sb.Append("                       ISNULL(T1.计秖,0) ╰参计秖");
            sb.Append("                         FROM WH_PACK T0");
            sb.Append("                       LEFT JOIN (SELECT MODEL_NO,MODEL_VER,MAX(PAL_QTY) 计秖 FROM CART ");
            sb.Append("                       GROUP BY MODEL_NO,MODEL_VER) T1 ON (Substring ([ItemCode],2,8) =MODEL_NO AND Substring([ItemCode],12,1)=MODEL_VER )");
            sb.Append(" where SUBSTRING(ITEMCODE,1,1) <> 'O'");
            if (textBox1.Text != "")
            {
                sb.Append(" AND INVOICE  like '%" + textBox1.Text + "%'  ");
            }
            sb.Append(" UNION ALL");
            sb.Append("                       SELECT SUBSTRING(ITEMCODE,1,9) MODEL,Substring([ItemCode],12,1) VER,DOCDATE DATE,INVOICE 'AU INV#',ITEMCODE 'ACME P/N',cast(CASE PACK WHEN '' THEN 0 ELSE PACK END as decimal)  '狾计',cast(ISNULL(CASE QTY WHEN '' THEN 0 ELSE QTY END,0) as decimal) 计秖");
            sb.Append("                       ,CASE WHEN (ISNULL(T1.计秖,0)-CASE WHEN CAST((CASE QTY WHEN '' THEN 0 ELSE QTY END) AS DECIMAL)=0 THEN 0");
            sb.Append("                       WHEN CAST(CASE PACK WHEN '' THEN 0 ELSE PACK END AS DECIMAL)=0 THEN 0 ELSE CAST((CASE QTY WHEN '' THEN 0 ELSE QTY END) AS DECIMAL)/CAST(CASE PACK WHEN '' THEN 0 ELSE PACK END AS DECIMAL) END ) <> 0 THEN '钵盽' END '钵盽',");
            sb.Append("                       ISNULL(T1.计秖,0) ╰参计秖");
            sb.Append("                         FROM WH_PACK T0");
            sb.Append("                       LEFT JOIN (SELECT MODEL_NO,MODEL_VER,MAX(PAL_QTY) 计秖 FROM CART ");
            sb.Append("                       GROUP BY MODEL_NO,MODEL_VER) T1 ON (Substring ([ItemCode],2,8) =MODEL_NO AND Substring([ItemCode],12,1)=MODEL_VER )");
            sb.Append(" where SUBSTRING(ITEMCODE,1,1) = 'O'   ");
            if (textBox1.Text != "")
            {
                sb.Append(" AND INVOICE  like '%" + textBox1.Text + "%'  ");
            }
            sb.Append(" ) AS A  ");
        

            sb.Append("  ORDER BY MODEL,VER");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Get4(string MODEL_NO, string MODEL_VER)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT cast(ISNULL(CASE PAL_QTY WHEN '' THEN 0 ELSE PAL_QTY END,0) as decimal)  计秖 FROM CART WHERE CASE WHEN substring(model_no,1,2) IN ('OM','OT') THEN 'O'+SUBSTRING((CASE CHARINDEX('(', MODEL_NO) WHEN 0 THEN  MODEL_NO ELSE SUBSTRING(MODEL_NO,0,CHARINDEX('(', MODEL_NO)) END),3,12) ELSE (CASE CHARINDEX('(', MODEL_NO) WHEN 0 THEN  MODEL_NO ELSE SUBSTRING(MODEL_NO,0,CHARINDEX('(', MODEL_NO)) END) END = @MODEL_NO AND MODEL_VER=@MODEL_VER ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL_NO", MODEL_NO));
            command.Parameters.Add(new SqlParameter("@MODEL_VER", MODEL_VER));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Get44(string MODEL_NO, string MODEL_VER)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT cast(ISNULL(CASE PAL_QTY WHEN '' THEN 0 ELSE PAL_QTY END,0) as decimal)  计秖 FROM CART");
            sb.Append("  WHERE CASE CHARINDEX('(', MODEL_NO) ");
            sb.Append(" WHEN 0 THEN  MODEL_NO ");
            sb.Append(" ELSE SUBSTRING(MODEL_NO,0,CHARINDEX('(', MODEL_NO)) END  = @MODEL_NO AND MODEL_VER=@MODEL_VER ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL_NO", MODEL_NO));
            command.Parameters.Add(new SqlParameter("@MODEL_VER", MODEL_VER));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("DATE", typeof(string));
            dt.Columns.Add("AU INV#", typeof(string));
            dt.Columns.Add("ACME P/N", typeof(string));
            dt.Columns.Add("狾计", typeof(string));
            dt.Columns.Add("计秖", typeof(string));
            dt.Columns.Add("╰参计秖", typeof(string));
 

            return dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                double D1 = 0;
          
                System.Data.DataTable dtCost = MakeTableCombine();
                System.Data.DataTable dt = Get3();
                System.Data.DataTable dt1 = null;
                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    DataRow dd = dt.Rows[i];
                    dr = dtCost.NewRow();
                    dr["DATE"] = dd["DATE"].ToString();
                    dr["AU INV#"] = dd["AU INV#"].ToString();

                    string f1 = dd["ACME P/N"].ToString();
                    dr["ACME P/N"] = f1;
                    string MODEL = dd["MODEL"].ToString();
                    string VER = dd["VER"].ToString();
                    dr["狾计"] = dd["狾计"].ToString();
                    dr["计秖"] = dd["计秖"].ToString();
                    D1 = Convert.ToDouble(dd["计秖"].ToString());
                    if (MODEL == "B125XTN02")
                    {
                        MessageBox.Show("a");
                    }
                    dt1 = Get4(MODEL, VER);
                    if(dt1.Rows.Count ==0)
                    {

                        dt1 = Get44(MODEL, VER);
                    }
                    StringBuilder sb = new StringBuilder();
                    for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                    {
                        DataRow dv = dt1.Rows[j];
                        string GH = dv["计秖"].ToString();
                        double D2 = D1 / Convert.ToDouble(dv["计秖"].ToString());
                        double D3 = Math.Ceiling(D2);
                        string D4 = D3.ToString();
                        if (!String.IsNullOrEmpty(GH))
                        {
                            if (GH != "0")
                            {
                                sb.Append(GH + "=" + D4 + "狾" + "/");
                            }

                        }
                    }
                    if (!String.IsNullOrEmpty(sb.ToString()))
                    {
                        sb.Remove(sb.Length - 1, 1);
                    }

                    string D = sb.ToString();
                    int S = D.IndexOf("礚絘");
                    if (S != -1)
                    {
                        D = "";
                    }
                    dr["╰参计秖"] = D;
                    dtCost.Rows.Add(dr);
                }
                //dataGridView1.DataSource = Get3();
                dataGridView1.DataSource = dtCost;

                label1.Text = "狾计 " + Get3().Compute("Sum(狾计)", null).ToString();
                label2.Text = "计秖 " + Get3().Compute("Sum(计秖)", null).ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void StacPack_Load(object sender, EventArgs e)
        {
            label1.Text = "";
            label2.Text = "";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //try
            //{
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("叫匡拒郎");
                }
                else
                {

                    GetExcelWH_SHIP(opdf.FileName);

     

                }

            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }



        private void GetExcelWH_SHIP(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
               int s = excelBook.Sheets.Count;

       
                   Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
                   string d = excelSheet.Name.Trim().ToString();
                   int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

                   int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

                   Microsoft.Office.Interop.Excel.Range range = null;


                   try
                   {


                       string id;
                       string id2 = "";
                       string id3 = "";
                       string id4 = "";
                       int u = 0;




                       for (int i = 1; i <= iRowCnt; i++)
                       {


                           range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                           range.Select();
                           string jh = range.Text.ToString();
                           int S = jh.IndexOf("ó");
                           if (S == -1)
                           {
                               range.EntireRow.Delete(XlDirection.xlUp);


                           }

                           //if (S != -1)
                           //{
                           //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                           //    string FG = range.Text.ToString();
                           //    int F = FG.IndexOf("");
                           //    id4 = FG.Substring(F+1, 8);

                           //}

                           //  int f = jh.IndexOf("ó");
                           //if (f != -1)
                           //{
                           //    id = jh.Substring(f + 3, jh.Length - f - 3);
                           //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 12, 1]);
                           //    string dd = range.Text.ToString();
                           //    id2 = dd.Substring(2, dd.Length - 2);

                           //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i + 14, 1]);
                           //    id3 = range.Text.ToString();
                           //    AddWH_SHIP(id4, id, id3, id2);
                           //}




                       }
                   }
                   finally
                   {
                      string  NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
           "Acme_" + Path.GetFileNameWithoutExtension(ExcelFile) + ".xls";


                       try
                       {
                           excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                       }
                       catch
                       {
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
                       MessageBox.Show("蹲Θ");
                   }
        }


        public void AddWH_SHIP(string DOCDATE, string DOCTYPE, string CARDNAME, string LOCATION)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_SHIP(INSDATE,DOCDATE,DOCTYPE,CARDNAME,LOCATION) values(@INSDATE,@DOCDATE,@DOCTYPE,@CARDNAME,@LOCATION)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@INSDATE", DateTime.Now.ToString("yyyyMMdd")));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            

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

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = GetTYPE();
        }

        private System.Data.DataTable GetTYPE()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  CASE WHEN month(DOCDATE) BETWEEN 1 AND 3 THEN 1 ");
            sb.Append("   WHEN month(DOCDATE) BETWEEN 4 AND 6 THEN 2");
            sb.Append("  WHEN month(DOCDATE) BETWEEN 7 AND 9 THEN 3");
            sb.Append("  WHEN month(DOCDATE) BETWEEN 10 AND 12 THEN 4 ");
            sb.Append(" END ﹗,DOCTYPE ó,");
            if (checkBox1.Checked)
            {
                sb.Append(" CARDNAME め,");
            }
            sb.Append("    COUNT(*) Ω计  FROM  WH_SHIP");
            sb.Append(" GROUP BY CASE WHEN month(DOCDATE) BETWEEN 1 AND 3 THEN 1 ");
            sb.Append("   WHEN month(DOCDATE) BETWEEN 4 AND 6 THEN 2");
            sb.Append("  WHEN month(DOCDATE) BETWEEN 7 AND 9 THEN 3");
            sb.Append("  WHEN month(DOCDATE) BETWEEN 10 AND 12 THEN 4 ");
            sb.Append(" END,DOCTYPE  ");
            if (checkBox1.Checked)
            {
                sb.Append(" ,CARDNAME ");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "ladingm ");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }



        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

       

    

    }
}
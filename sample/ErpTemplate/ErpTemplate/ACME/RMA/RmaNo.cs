using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
namespace ACME
{
    public partial class RmaNo : Form
    {
        string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql05;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public string q = "";
        public string q1;
        public string q2 = "";
        public RmaNo()
        {
            InitializeComponent();
        }

        private void ViewBatchPayment4()
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            if (textBox1.Text.Length + textBox2.Text.Length <= 3)
            {
                sb.Append(" select '進金生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql02.DBO.octr WHERE 1=2 ");
            }
            else
            {
                if (q1 == "進金生")
                {
                    sb.Append(" select '進金生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql02.DBO.octr WHERE 1=1 ");
                    if (textBox1.Text != "")
                    {
                        sb.Append(" and U_RMA_NO like  '%" + textBox1.Text.ToString().Trim() + "%'  ");
                    }

                    if (textBox2.Text != "")
                    {
                        sb.Append(" and U_AUO_RMA_NO like  '%" + textBox2.Text.ToString().Trim() + "%'  ");
                    }
                }
                if (q1 == "達睿生")
                {
                    sb.Append(" select '達睿生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql05.DBO.octr WHERE 1=1 ");
                    if (textBox1.Text != "")
                    {
                        sb.Append(" and U_RMA_NO like  '%" + textBox1.Text.ToString().Trim() + "%'  ");
                    }

                    if (textBox2.Text != "")
                    {
                        sb.Append(" and U_AUO_RMA_NO like  '%" + textBox2.Text.ToString().Trim() + "%'  ");
                    }
                }
                if (q1 == "進金生達睿生")
                {
                    sb.Append(" select '進金生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql02.DBO.octr WHERE 1=1 ");
                    if (textBox1.Text != "")
                    {
                        sb.Append(" and U_RMA_NO like  '%" + textBox1.Text.ToString().Trim() + "%'  ");
                    }

                    if (textBox2.Text != "")
                    {
                        sb.Append(" and U_AUO_RMA_NO like  '%" + textBox2.Text.ToString().Trim() + "%'  ");
                    }
                    sb.Append(" UNION ALL");
                    sb.Append(" select '達睿生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql05.DBO.octr WHERE 1=1 ");
                    if (textBox1.Text != "")
                    {
                        sb.Append(" and U_RMA_NO like  '%" + textBox1.Text.ToString().Trim() + "%'  ");
                    }

                    if (textBox2.Text != "")
                    {
                        sb.Append(" and U_AUO_RMA_NO like  '%" + textBox2.Text.ToString().Trim() + "%'  ");
                    }
                }
            }
     
            sb.Append(" order by u_rma_no desc  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            //填入精靈名稱


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "INV1");
            }
            finally
            {
                connection.Close();
            }


            bindingSource1.DataSource = ds.Tables[0];
            dataGridView1.DataSource = bindingSource1;

        }


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;


                    for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                    {

                        row = dataGridView1.SelectedRows[i];

                        string COMOANY = row.Cells["COMPANY"].Value.ToString();
                        if (COMOANY == "進金生")
                        {
                            listBox1.Items.Add(row.Cells["Contractid"].Value.ToString());
                        }
                        if (COMOANY == "達睿生")
                        {
                            listBox2.Items.Add(row.Cells["Contractid"].Value.ToString());
                        }
                    }

                    if (listBox1.Items.Count > 0)
                    {
                        ArrayList al = new ArrayList();

                        for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                        {
                            al.Add(listBox1.Items[i].ToString());
                        }
                        StringBuilder sb = new StringBuilder();
                        foreach (string v in al)
                        {
                            sb.Append("'" + v + "',");
                        }

                        sb.Remove(sb.Length - 1, 1);
                        q = sb.ToString();
                    }

                     if (listBox2.Items.Count > 0)
                     {
                         ArrayList al2 = new ArrayList();

                         for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                         {
                             al2.Add(listBox2.Items[i].ToString());
                         }
                         StringBuilder sb2 = new StringBuilder();
                         foreach (string v2 in al2)
                         {
                             sb2.Append("'" + v2 + "',");
                         }

                         sb2.Remove(sb2.Length - 1, 1);
                         q2 = sb2.ToString();
                     }

                }
                else
                {
                    MessageBox.Show("請點選單號");
                    return;
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

       

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            ViewBatchPayment4();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ViewBatchPayment4();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("請選擇檔案");
                }
                else
                {

                    GD4(opdf.FileName);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public  System.Data.DataTable GetOCTR(string U_RMA_NO,string  F1)
        {
            SqlConnection MyConnection = globals.shipConnection;
            if (F1 == "D")
            {
                MyConnection = new SqlConnection(strCn02);
            }
            string sql = "SELECT U_AUO_RMA_NO,CASE CHARINDEX('_', U_RMODEL) WHEN 0 THEN U_RMODEL ELSE   SUBSTRING(U_RMODEL,0,CHARINDEX('_', U_RMODEL)) END U_RMODEL,U_RVER FROM OCTR  WHERE U_RMA_NO=@U_RMA_NO ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " octr ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" octr "];
        }
        private void GD4(string ExcelFile)
        {

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

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            string RMANO;
            string VENRMANO;
            string MODEL;
            string VER;
            StringBuilder sb = new StringBuilder();

            for (int i = 2; i <= iRowCnt; i++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                range.Select();
                RMANO = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
            //    range.NumberFormatLocal = "0_ ";
                VENRMANO = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                MODEL = range.Text.ToString().Trim().ToUpper().Replace("_OPEN CELL", "");

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                VER = range.Text.ToString().Trim();

                if (!String.IsNullOrEmpty(RMANO))
                {
                    try
                    {
                        string F1 = RMANO.Substring(0, 1);
                        int H1 = 0;
                        System.Data.DataTable G1 = GetOCTR(RMANO, F1);
                        string U_AUO_RMA_NO = G1.Rows[0]["U_AUO_RMA_NO"].ToString();
                        string U_RMODEL = G1.Rows[0]["U_RMODEL"].ToString();
                        string U_RVER = G1.Rows[0]["U_RVER"].ToString();
                        if (G1.Rows.Count == 0)
                        {
                            H1 = 1;
                            MessageBox.Show("SAP沒有RMA NO " + RMANO + " 資料");
                        }
                        if (VENRMANO != U_AUO_RMA_NO)
                        {
                            if (H1 == 0)
                            {
                                H1 = 1;
                                MessageBox.Show("RMANO " + RMANO + " VENDER NO錯誤");
                            }
                        }
                        if (U_RMODEL != MODEL)
                        {
                            if (H1 == 0)
                            {
                                H1 = 1;
                                MessageBox.Show("RMANO " + RMANO + " MODEL錯誤");
                            }
                        }
                        if (U_RVER != VER)
                        {
                            if (H1 == 0)
                            {
                                H1 = 1;
                                MessageBox.Show("RMANO " + RMANO + " VER錯誤");
                            }
                        }

                        if (H1 == 0)
                        {
                            listBox3.Items.Add(VENRMANO + ' ' + RMANO);
                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }

            if (listBox3.Items.Count > 0)
            {
                ArrayList al = new ArrayList();

                for (int f = 0; f <= listBox3.Items.Count - 1; f++)
                {
                    al.Add(listBox3.Items[f].ToString());
                }
     
                foreach (string v in al)
                {
                    sb.Append("'" + v + "',");
                }

                sb.Remove(sb.Length - 1, 1);

            }

    //        System.Diagnostics.Process[] ps =
    //System.Diagnostics.Process.GetProcessesByName("EXCEL");

    //        foreach (System.Diagnostics.Process p in ps)
    //        {
    //            p.Kill();
    //        }
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

            if (!String.IsNullOrEmpty(sb.ToString()))
            {
                SqlConnection connection = globals.shipConnection;

                StringBuilder sb2 = new StringBuilder();


                if (q1 == "進金生")
                {
                    sb2.Append(" select '進金生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql02.DBO.octr WHERE 1=1 ");
                    sb2.Append(" and U_AUO_RMA_NO+' '+U_RMA_NO in  (" + sb.ToString() + ")  ");
                }
                if (q1 == "達睿生")
                {
                    sb2.Append(" select '達睿生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql05.DBO.octr WHERE 1=1 ");
                    sb2.Append(" and U_AUO_RMA_NO+' '+U_RMA_NO in  (" + sb.ToString() + ")  ");
                }
                if (q1 == "進金生達睿生")
                {
                    sb2.Append(" select '進金生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql02.DBO.octr WHERE 1=1 ");
                    sb2.Append(" and U_AUO_RMA_NO+' '+U_RMA_NO in  (" + sb.ToString() + ")  ");
                    sb2.Append(" UNION ALL");
                    sb2.Append(" select '達睿生' COMPANY,U_RMA_NO,U_CUSNAME_S,U_RMODEL,U_RVER,U_RGRADE,U_RQUINITY,Contractid,U_AUO_RMA_NO   from acmesql05.DBO.octr WHERE 1=1 ");
                    sb2.Append(" and U_AUO_RMA_NO+' '+U_RMA_NO in  (" + sb.ToString() + ")  ");

                }

                sb2.Append(" order by u_rma_no desc  ");


                SqlCommand command = new SqlCommand(sb2.ToString(), connection);
                command.CommandType = CommandType.Text;

                //填入精靈名稱


                SqlDataAdapter da = new SqlDataAdapter(command);

                DataSet ds = new DataSet();
                try
                {
                    connection.Open();
                    da.Fill(ds, "INV1");
                }
                finally
                {
                    connection.Close();
                }


                bindingSource1.DataSource = ds.Tables[0];
                dataGridView1.DataSource = bindingSource1;
            }
        }
     
    }

}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
namespace ACME
{
    public partial class ADAUO : Form
    {

        public ADAUO()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetSGROUP("進貨");
            dataGridView2.DataSource = GetSGROUP("出貨");
            System.Data.DataTable GG1 = GetSGROUP2();
            if (GG1.Rows.Count > 0)
            {
                dataGridView3.DataSource = GG1;
                dataGridView3.Visible = true;
            }
            else
            {
                dataGridView3.Visible = false;
            }
        }

        private System.Data.DataTable GetSGROUP(string DOCTYPE)
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CARDNAME 客戶,CHINO 正航單號,DOCDATE 日期,INVOICENO,WT Watt,MODEL 'Model Name'");
            sb.Append(" ,PALLET 'Pallet No',CARTON 'Carton No',SN 'Serial No',GLASS 'Glass ID',PM 'Pmax(W)',VM 'Vmpp(V)'");
            sb.Append(" ,AM 'Impp(A)',VOC 'Voc(V)',ISC 'Isc(A)',FF 'F.F.',CONT 'Container No',MEMO 備註,LOC 案場 FROM ACMESQLSP.DBO.AD_AUO WHERE DOCTYPE=@DOCTYPE");

            sb.Append(" and  DOCDATE between @DOCDATE1 and @DOCDATE2 ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  CHINO between @CHINO1 and @CHINO2 ");
            }
            if (textBox5.Text != "")
            {
                sb.Append(" and  CARDNAME  LIKE '%" + textBox5.Text + "%' ");
            }
            if (textBox13.Text != "")
            {
                sb.Append(" and  INVOICENO =@INVOICENO ");
            }
            if (textBox7.Text != "")
            {
                sb.Append(" and  MODEL =@MODEL ");
            }
            if (textBox8.Text != "")
            {
                sb.Append(" and  WT =@WT ");
            }
            if (textBox9.Text != "")
            {
                sb.Append(" and  PALLET =@PALLET ");
            }
            if (textBox10.Text != "")
            {
                sb.Append(" and  CARTON =@CARTON ");
            }
            if (textBox11.Text != "")
            {
                sb.Append(" and  SN =@SN ");
            }
            if (textBox12.Text != "")
            {
                sb.Append(" and  GLASS =@GLASS ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@DOCDATE1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DOCDATE2", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@CHINO1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CHINO2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@INVOICENO", textBox13.Text));
            command.Parameters.Add(new SqlParameter("@MODEL", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@WT", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@PALLET", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@CARTON", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@SN", textBox11.Text));
            command.Parameters.Add(new SqlParameter("@GLASS", textBox12.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }
        private System.Data.DataTable GetSGROUP2()
        {
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DOCTYPE ' ',COUNT(*) 數量  FROM ACMESQLSP.DBO.AD_AUO  ");

            sb.Append(" WHERE  DOCDATE between @DOCDATE1 and @DOCDATE2 ");
            if (textBox1.Text != "" && textBox2.Text != "")
            {
                sb.Append(" and  CHINO between @CHINO1 and @CHINO2 ");
            }
            if (textBox5.Text != "")
            {
                sb.Append(" and  CARDNAME  LIKE '%" + textBox5.Text + "%' ");
            }
            if (textBox13.Text != "")
            {
                sb.Append(" and  INVOICENO =@INVOICENO ");
            }
            if (textBox7.Text != "")
            {
                sb.Append(" and  MODEL =@MODEL ");
            }
            if (textBox8.Text != "")
            {
                sb.Append(" and  WT =@WT ");
            }
            if (textBox9.Text != "")
            {
                sb.Append(" and  PALLET =@PALLET ");
            }
            if (textBox10.Text != "")
            {
                sb.Append(" and  CARTON =@CARTON ");
            }
            if (textBox11.Text != "")
            {
                sb.Append(" and  SN =@SN ");
            }
            if (textBox12.Text != "")
            {
                sb.Append(" and  GLASS =@GLASS ");
            }
            sb.Append("            GROUP BY DOCTYPE ");
 
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCDATE1", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@DOCDATE2", textBox4.Text));
            command.Parameters.Add(new SqlParameter("@CHINO1", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@CHINO2", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@INVOICENO", textBox13.Text));
            command.Parameters.Add(new SqlParameter("@MODEL", textBox7.Text));
            command.Parameters.Add(new SqlParameter("@WT", textBox8.Text));
            command.Parameters.Add(new SqlParameter("@PALLET", textBox9.Text));
            command.Parameters.Add(new SqlParameter("@CARTON", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@SN", textBox11.Text));
            command.Parameters.Add(new SqlParameter("@GLASS", textBox12.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                connection.Close();
            }


            return ds.Tables[0];


        }

        private void button2_Click(object sender, EventArgs e)
        {

            if (tabControl1.SelectedTab == 進貨)
            {
                ExcelReport.GridViewToExcel(dataGridView1);
            }
            else if (tabControl1.SelectedTab == 出貨)
            {
                ExcelReport.GridViewToExcel(dataGridView2);
            }
        }

        private void HROVER_Load(object sender, EventArgs e)
        {
            textBox3.Text = DateTime.Now.ToString("yyyy") + "0101";
            textBox4.Text = GetMenu.DLast();

            System.Data.DataTable G1 = GetMenu.Getdata("ADAUO");
            if (G1.Rows.Count > 0)
            {
                textBox6.Text = G1.Rows[0][0].ToString();
            }
        }



        public void AddAUO(string DOCTYPE, string CHINO, string CARDNAME, string DOCDATE, string INVOICENO, string WT, string MODEL, string PALLET, string CARTON, string SN, string GLASS, string PM, string VM, string AM, string VOC, string ISC, string FF, string CONT, string MEMO, string FILENAME, string LOC)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ACMESQLSP.DBO.AD_AUO(DOCTYPE,CHINO,CARDNAME,DOCDATE,INVOICENO,WT,MODEL,PALLET,CARTON,SN,GLASS,PM,VM,AM,VOC,ISC,FF,CONT,MEMO,FILENAME,LOC) values(@DOCTYPE,@CHINO,@CARDNAME,@DOCDATE,@INVOICENO,@WT,@MODEL,@PALLET,@CARTON,@SN,@GLASS,@PM,@VM,@AM,@VOC,@ISC,@FF,@CONT,@MEMO,@FILENAME,@LOC)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@CHINO", CHINO));
            command.Parameters.Add(new SqlParameter("@CARDNAME", CARDNAME));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@INVOICENO", INVOICENO));
            command.Parameters.Add(new SqlParameter("@WT", WT));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@PALLET", PALLET));
            command.Parameters.Add(new SqlParameter("@CARTON", CARTON));
            command.Parameters.Add(new SqlParameter("@SN", SN));
            command.Parameters.Add(new SqlParameter("@GLASS", GLASS));
            command.Parameters.Add(new SqlParameter("@PM", PM));
            command.Parameters.Add(new SqlParameter("@VM", VM));
            command.Parameters.Add(new SqlParameter("@AM", AM));
            command.Parameters.Add(new SqlParameter("@VOC", VOC));
            command.Parameters.Add(new SqlParameter("@ISC", ISC));
            command.Parameters.Add(new SqlParameter("@FF", FF));
            command.Parameters.Add(new SqlParameter("@CONT", CONT));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@FILENAME", FILENAME));
            command.Parameters.Add(new SqlParameter("@LOC", LOC));
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

        public void DELAUO(string  FILENAME)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE ACMESQLSP.DBO.AD_AUO WHERE FILENAME=@FILENAME", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@FILENAME", FILENAME));

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
        public void DELAUO2()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE ACMESQLSP.DBO.AD_AUO WHERE FILENAME  LIKE '%" + textBox14.Text + "%' ", connection);
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
        private void WriteAUO(string ExcelFile)
        {
            //  AddAP
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                DELAUO(excelFile);
                string CARDNAME;
                string MODEL;
                string PALLET;
                string CARTON;
                string SN;
                string GLASS;
                string PM;
                string VM;
                string AM;
                string VOC;
                string ISC;
                string FF;
                string CONT;
                string CHINO;
                string INVOICENO;
                string WT;
                string DOCDATE;
                string MEMO;
                string LOC;
                string DOCTYPE = "";
                int G1 = excelFile.IndexOf("進貨");
                int G2 = excelFile.IndexOf("出貨");
                if (G1 != -1)
                {
                    DOCTYPE = "進貨";
                }
                if (G2 != -1)
                {
                    DOCTYPE = "出貨";
                }
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    CARDNAME = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    CHINO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    DOCDATE = range.Text.ToString().Trim().Replace("-", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    INVOICENO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    WT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    MODEL = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    PALLET = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    CARTON = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    SN = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    GLASS = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    PM = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 12]);
                    range.Select();
                    VM = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    AM = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    VOC = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 15]);
                    range.Select();
                    ISC = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 16]);
                    range.Select();
                    FF = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    CONT = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 18]);
                    range.Select();
                    MEMO = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 19]);
                    range.Select();
                    LOC = range.Text.ToString().Trim();

                    //Model Name
                    if (MODEL != "Model Name")
                    {
                        if (!String.IsNullOrEmpty(PALLET))
                        {
                            AddAUO(DOCTYPE, CHINO, CARDNAME, DOCDATE, INVOICENO, WT, MODEL, PALLET, CARTON, SN, GLASS, PM, VM, AM, VOC, ISC, FF, CONT, MEMO, excelFile, LOC);
                        }
                    }



                }




            }
            finally
            {


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


                //         System.Diagnostics.Process.Start(NewFileName);


            }



        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == 進貨)
            {

                for (int i = 0; i <= dataGridView1.Rows.Count - 2; i++)
                {

                    DataGridViewRow row;

                    row = dataGridView1.Rows[i];
                    string ID = row.Cells["ID"].Value.ToString();
                    string 備註 = row.Cells["備註"].Value.ToString();

                    if (!String.IsNullOrEmpty(備註))
                    {
                        UpdateMasterSQL(備註, ID);
                    }
                }
            }
            else if (tabControl1.SelectedTab == 出貨)
            {
                for (int i = 0; i <= dataGridView2.Rows.Count -2; i++)
                {

                    DataGridViewRow row;

                    row = dataGridView2.Rows[i];
                    string ID = row.Cells["ID"].Value.ToString();
                    string 備註 = row.Cells["備註"].Value.ToString();

                    if (!String.IsNullOrEmpty(備註))
                    {
                        UpdateMasterSQL(備註, ID);
                    }
                }
            }


            button1_Click(null, new EventArgs());
        }

        private void UpdateMasterSQL(string MEMO, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE  ACMESQLSP.DBO.AD_AUO SET MEMO=@MEMO WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@ID", ID));

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

        private void button5_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                string t1 = folderBrowserDialog1.SelectedPath.ToString();

                if (GetMenu.Getdata("ADAUO").Rows.Count == 0)
                {
                    GetMenu.Add(t1, "ADAUO");
                }
                else
                {
                    GetMenu.UP(t1, "ADAUO");
                }

                textBox6.Text = t1;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {

            string d = textBox6.Text;


            if (!String.IsNullOrEmpty(d))
            {
                string[] filenames = Directory.GetFiles(d);
      
                foreach (string file in filenames)
                {

                    int G1 = file.IndexOf("進貨");
                    int G2 = file.IndexOf("出貨");
                    if (G1 != -1 || G2 != -1)
                    {
                        WriteAUO(file);
                        File.Delete(file);
                    }
             

          
                }
                MessageBox.Show("上傳成功" );
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DELAUO2();
            MessageBox.Show("檔案已刪除");
        }

     
    }
}

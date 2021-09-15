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
    public partial class Fee : Form
    {
        public Fee()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                if (opdf.FileName.ToString() == "")
                {
                    MessageBox.Show("請選擇檔案");
                }
                else
                {
                    string F = opdf.FileName;
                    int df = F.IndexOf("代");
                    //string dd = F.Substring(df - 4, 4);
                    if (df != -1)
                    {
                        GetExcelContentGD4(F, 40);
                    }
                    else
                    {
                        GetExcelContentGD44(F, 40);
                    }

                }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //}
        }

        private void GetExcelContentGD4(string ExcelFile, int Y)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
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
            string idG = "";
            int u = 0;
            int v = 0;
            int L1 = 0;
            for (int b = 5; b <= 20; b++)
            {
                for (int jj = 1; jj <= 20; jj++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[jj, b]);
                    range.Select();
                    id = range.Text.ToString().Trim();

                    int G1 = id.IndexOf("金额");

                    if (G1 != -1)
                    {
                 
                        u = jj;
                        v = b + 1 ;
                        break;
                    }

                }
            }

            for (int U = u + 2; U <= 1000; U++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[U, 1]);
                range.Select();
                id = range.Text.ToString().Trim();
                if (String.IsNullOrEmpty(id))
                {
                    L1 = U - 1;
                    break;
                }

            }
            if (u == 0)
            {
                MessageBox.Show("Excel格式有誤");
                return;

            }
            for (int i = v; i <= Y; i++)
            {




                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[u, i]);
                range.Select();
                id = range.Text.ToString().Trim();

          

                //try
                //{


                if (id != "车型" )
                    {



                        for (int j = u; j <= L1; j++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, i]);
                            range.Select();
                            id3 = range.Text.ToString().Trim();


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, 5]);
                            range.Select();

                            id4 = range.Text.ToString().Trim();



                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, 3]);
                            range.Select();
                            id2 = range.Text.ToString().Trim();

                            int FF = id4.IndexOf("SH");
                            int FF2 = id4.IndexOf("SI");
                            if (FF.ToString() != "-1" || FF2.ToString() != "-1")
                            {

                                if ((!String.IsNullOrEmpty(id4)) && id3.Trim() != "" && id3.Trim() != "0" && id3.Trim() != "0.00" && id3.Trim() != "/")
                                {
                                    string hj = "";
                                    if (comboBox2.Text != "蘇州宏高")
                                    {
                                        hj = "";
                                    }
                                    else
                                    {
                                        hj = comboBox3.Text;
                                    }

                                            decimal n;
                                            if (decimal.TryParse(id3, out n))
                                            {
                                                decimal cd = Convert.ToDecimal(id3) * Convert.ToDecimal(textBox1.Text);

                                                if (id.Trim() != "金额" && id.Trim().ToUpper() != "TOTAL" && !String.IsNullOrEmpty(id))
                                                {

                                                    if (fmLogin.LoginID.ToString().ToUpper() != "LLEYTONCHEN")
                                                    {
                                                        AddAUOGD4(id4, id, cd.ToString(), comboBox2.Text, comboBox3.Text, id2, comboBox1.Text, textBox1.Text, id3, comboBox6.Text);
                                                    }
                                                }
                                                else
                                                {
                                                    if (cd < 0)
                                                    {
                                                        if (fmLogin.LoginID.ToString().ToUpper() != "LLEYTONCHEN")
                                                        {
                                                            AddAUOGD4(id4, "", cd.ToString(), comboBox2.Text, comboBox3.Text, id2, comboBox1.Text, textBox1.Text, id3, comboBox6.Text);
                                                        }

                                                    }
                                                }
                                            }
                                }
                            }
                        

                    }


                }



                //}

                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.Message);
                //}
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
            MessageBox.Show("匯出成功");
        }


        private void GetExcelContentGD44(string ExcelFile,int Y)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

       //     int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

        //    Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string id2 = "";
            string id3 = "";
            string id4 = "";
            string idG = "";

            int u = 0;
            int v = 0;
            int L1 = 0;

   
            for (int b = 5; b <= 20; b++)
            {
                for (int jj = 1; jj <= 20; jj++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[jj, b]);
                    range.Select();
                    id = range.Text.ToString().Trim();




                    if (id == "金额")
                    {
                        u = jj + 1;
                        v = b + 1;
                        break;
                    }
                    
                }

            }
            for (int U = u+2; U <= 1000; U++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[U, 1]);
                range.Select();
                id = range.Text.ToString().Trim();
                if (String.IsNullOrEmpty(id))
                {
                    L1 = U - 1;
                    break;
                }

            }
        
            if (u == 0)
            {
                MessageBox.Show("Excel格式有誤");
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
                return;
            
            }
            for (int i = v; i <= Y; i++)
            {
            
                

             
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[u, i]);
                range.Select();
                id = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[u-1, i]);
                range.Select();
                idG = range.Text.ToString().Trim();

     

                //try
                //{


                if (id != "车型" )
                {



                    for (int j = u; j <= L1; j++)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, i]);
                                range.Select();
                                id3 = range.Text.ToString().Trim();


                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, 5]);
                                range.Select();

                                id4 = range.Text.ToString().Trim();



                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, 3]);
                                range.Select();
                                id2 = range.Text.ToString().Trim();

                                int FF = id4.IndexOf("SH");
                                int FF2 = id4.IndexOf("SI");
                                if (FF.ToString() != "-1" || FF2.ToString() != "-1")
                                {
                                    if ((!String.IsNullOrEmpty(id4)) && id3.Trim() != "" && id3.Trim() != "0" && id3.Trim() != "0.00" && id3.Trim() != "/")
                                    {
                                        string hj = "";
                                        if (comboBox2.Text != "蘇州宏高")
                                        {
                                            hj = "";
                                        }
                                        else
                                        {
                                            hj = comboBox3.Text;
                                        }
                                         decimal n;
                                         if (decimal.TryParse(id3, out n))
                                         {
                                             decimal cd = Convert.ToDecimal(id3) * Convert.ToDecimal(textBox1.Text);
                                             //if (cd == -1490)
                                             //{
                                             //    MessageBox.Show("A");
                                             //}
                                             if (idG.Trim() != "合计" && idG.Trim().ToUpper() != "TOTAL" && !String.IsNullOrEmpty(id))
                                             {
                                                 if (fmLogin.LoginID.ToString().ToUpper() != "LLEYTONCHEN")
                                                 {
                                                     AddAUOGD4(id4, id, cd.ToString(), comboBox2.Text, comboBox3.Text, id2, comboBox1.Text, textBox1.Text, id3,comboBox6.Text);
                                                 }
                                             }
                                             else
                                             {
                                                 if (cd < 0)
                                                 {
                                                     if (fmLogin.LoginID.ToString().ToUpper() != "LLEYTONCHEN")
                                                     {
                                                         AddAUOGD4(id4, "", cd.ToString(), comboBox2.Text, comboBox3.Text, id2, comboBox1.Text, textBox1.Text, id3, comboBox6.Text);
                                                     }
                                                 }
                                             }
                                         }
                                         
                                    }
                                }


                            }


                    }



                //}

                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.Message);
                //}
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
            MessageBox.Show("匯出成功");
        }

        public void AddAUOGD4(string shippingcode, string ITEM, string amount, string CardName, string SubCompany, string DocDate, string DocCur, string DocCur1, string amount2, string FMONTH)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into SHIPPING_FEE(shippingcode,ITEM,amount,CardName,SubCompany,DocDate,DocCur,DocCur1,InsDate,amount2,FMONTH) values(@shippingcode,@ITEM,@amount,@CardName,@SubCompany,@DocDate,@DocCur,@DocCur1,@InsDate,@amount2,@FMONTH)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@ITEM", ITEM));
            command.Parameters.Add(new SqlParameter("@amount", amount));
            command.Parameters.Add(new SqlParameter("@CardName", CardName));
            command.Parameters.Add(new SqlParameter("@SubCompany", SubCompany));
            command.Parameters.Add(new SqlParameter("@DocDate", DocDate));
            command.Parameters.Add(new SqlParameter("@DocCur", DocCur));
            command.Parameters.Add(new SqlParameter("@DocCur1", DocCur1));
            command.Parameters.Add(new SqlParameter("@InsDate", DateTime.Now.ToString("yyyyMMdd") ));
            command.Parameters.Add(new SqlParameter("@amount2", amount2));
            command.Parameters.Add(new SqlParameter("@FMONTH", FMONTH));
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
            SqlCommand command = new SqlCommand("TRUNCATE TABLE SHIPPING_FEE ", connection);
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
        public void DELETE(string FF)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE SHIPPING_FEE WHERE ID IN (" + FF + ")", connection);
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


        public void UPDATESAP(string SAP, string aa, string bb, string cc)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE SHIPPING_FEE SET SAP=@SAP  where InsDate  between @aa and @bb and isnull(feecheck,'False')=@cc ", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@SAP", SAP));
            command.Parameters.Add(new SqlParameter("@aa", aa));
            command.Parameters.Add(new SqlParameter("@bb", bb));
            command.Parameters.Add(new SqlParameter("@cc", cc));


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
        private void Fee_Load(object sender, EventArgs e)
        {
            UtilSimple.SetLookupBinding(comboBox7, GetMenu.Month2(), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox6, GetMenu.Month2(), "DataValue", "DataValue");
           
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.Getfee("feeDoccur"), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.Getfee("feeCardName"), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox3, GetMenu.Getfee("feeSubName"), "DataValue", "DataValue");
            UtilSimple.SetLookupBinding(comboBox4, GetMenu.Getfee("JOYCHECK"), "DataText", "DataValue");
            UtilSimple.SetLookupBinding(comboBox5, GetMenu.Getfee2("feeCardName"), "DataValue", "DataValue");
            comboBox2.Text = "蘇州宏高";
            comboBox3.Text = "蘇州";
            
            comboBox1.Text = "NTD";
            comboBox5.Text = "";

            textBox2.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox3.Text = DateTime.Now.ToString("yyyyMMdd");

            textBox5.Text = DateTime.Now.ToString("yyyyMMdd");
            textBox6.Text = DateTime.Now.ToString("yyyyMMdd");


        }



        private void button3_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.shipping_Fee1BindingSource.EndEdit();
            this.shipping_Fee1TableAdapter.Update(this.ship.Shipping_Fee1);

            MessageBox.Show("存檔成功");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string FF = comboBox7.Text;
                this.shipping_Fee1TableAdapter.Fill(this.ship.Shipping_Fee1, textBox2.Text, textBox3.Text, comboBox4.SelectedValue.ToString(), comboBox5.SelectedValue.ToString(), comboBox7.Text);

                dataGridView1.DataSource = Get1();

                dataGridView2.DataSource = Get2();


                checkBox1.Checked = false;

                //for (int i = 0; i < this.shipping_Fee1DataGridView.Rows.Count; i++)
                //{
                //    this.shipping_Fee1DataGridView.Rows[i].Cells[11].Value = false;
                //}
                for (int i = 0; i < this.shipping_Fee1DataGridView.Rows.Count; i++)
                {
                    this.shipping_Fee1DataGridView.Rows[i].Cells[12].Value = false;

                    //MessageBox.Show(shipping_Fee1DataGridView.Rows[i].Cells[11].Value.ToString());
                    if (shipping_Fee1DataGridView.Rows[i].Cells[11].Value.ToString().Trim() == "True")
                    {
                        shipping_Fee1DataGridView.Rows[i].Cells[11].ReadOnly = true;
                    }
                }

              
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }



        private void shipping_Fee1DataGridView_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            if (shipping_Fee1DataGridView.SelectedRows.Count > 0)
            {

                string da = shipping_Fee1DataGridView.SelectedRows[0].Cells["ShippingCode"].Value.ToString();

                fmShip a = new fmShip();
                a.PublicString = da;
                //a.WindowState = FormWindowState.Normal;
                //a.StartPosition = FormStartPosition.CenterScreen;

                //a.MdiParent = null;
                a.ShowDialog();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcelSelectJOY(shipping_Fee1DataGridView);
        }


        private System.Data.DataTable Get1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT boardCountNo 貿易條件,SUM(CAST(T0.AMOUNT AS DECIMAL(15,5))) 金額 FROM dbo.Shipping_Fee T0");
            sb.Append("   LEFT JOIN SHIPPING_MAIN T1 ON (RTRIM(LTRIM(T0.SHIPPINGCODE))=T1.SHIPPINGCODE)");
            sb.Append(" WHERE  T0.INSDATE BETWEEN @AA AND @BB and  isnull(feecheck,'False')=@CC  ");
            if (comboBox5.Text != "")
            {
                sb.Append(" AND t0.cardname =@cardname ");
            }
            if (comboBox7.Text != "")
            {
                sb.Append(" AND t0.FMONTH =@FMONTH ");
            }
            sb.Append(" GROUP BY boardCountNo");
         

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@AA", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC", comboBox4.SelectedValue.ToString()));
            command.Parameters.Add(new SqlParameter("@cardname", comboBox5.SelectedValue.ToString()));
            command.Parameters.Add(new SqlParameter("@FMONTH", comboBox7.Text));
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
        private System.Data.DataTable Get2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUBSTRING(T3.GROUPNAME,3,15) BU,t1.boardCountNo 貿易條件,t0.ShippingCode 工單號碼,item 費用,amount 金額,t0.cardname 供應商,subcompany 子公司,DocDate 日期,doccur 幣別,doccur1 匯率 FROM dbo.Shipping_Fee T0 ");
            sb.Append("   LEFT JOIN SHIPPING_MAIN T1 ON (RTRIM(LTRIM(T0.SHIPPINGCODE))=T1.SHIPPINGCODE)");
            sb.Append(" left join acmesql02.dbo.ocrd t2 on (t2.cardcode=t1.cardcode COLLATE Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" LEFT JOIN acmesql02.dbo.OCRG T3 ON (T2.GROUPCODE = T3.GROUPCODE) ");
            sb.Append(" WHERE  T0.INSDATE BETWEEN @AA AND @BB and  isnull(feecheck,'False')=@CC ");
            if (comboBox5.Text != "")
            {
                sb.Append(" AND T0.cardname =@cardname ");
            }

            if (comboBox7.Text != "")
            {
                sb.Append(" AND T0.FMONTH =@FMONTH ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@AA", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox3.Text));
            command.Parameters.Add(new SqlParameter("@CC", comboBox4.SelectedValue.ToString()));
            command.Parameters.Add(new SqlParameter("@cardname", comboBox5.SelectedValue.ToString()));
            command.Parameters.Add(new SqlParameter("@FMONTH", comboBox7.Text));
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

        private System.Data.DataTable Get3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT InsDate 匯入日期,ItemType 類型,ItemDate 更動時間,ShippingCode 工單號碼,Item 費用名稱,FeeCheck 已審核,Amount 金額,CardName 供應商  FROM dbo.Shipping_Fee_LOG T0 ");
            sb.Append(" WHERE  T0.INSDATE BETWEEN @AA AND @BB  ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;


            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@AA", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox3.Text));
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

        private void button5_Click(object sender, EventArgs e)
        {

         //   checkBox1.Checked = false;

          

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i <= shipping_Fee1DataGridView.Rows.Count - 1; i++)
            {

                DataGridViewRow row;
             
                row = shipping_Fee1DataGridView.Rows[i];
                if (!String.IsNullOrEmpty(row.Cells["aaa"].Value.ToString()))
                {
                    string a0 = row.Cells["aaa"].Value.ToString();
 
                    if (a0 == "True")
                    {
                      //  string a1 = row.Cells["FeeCheck"].Value.ToString();
                        string ID = row.Cells["ID"].Value.ToString();

                        //if (a1.Trim() != "True")
                        //{
           
                            sb.Append("'" + ID + "',");
                       //}

                      
                    }
             
             
                }
              
            }
            if (!String.IsNullOrEmpty(sb.ToString()))
            {
                sb.Remove(sb.Length - 1, 1);
                DELETE(sb.ToString());
                MessageBox.Show("資料已刪除");
                checkBox1.Checked = false;
                button2_Click(null, new EventArgs());
 
            }
        
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void shipping_Fee1DataGridView_MouseDoubleClick_1(object sender, MouseEventArgs e)
        {
            if (shipping_Fee1DataGridView.SelectedRows.Count > 0)
            {

                string da = shipping_Fee1DataGridView.SelectedRows[0].Cells["ShippingCode"].Value.ToString();

                fmShip a = new fmShip();
                a.PublicString = da;
       
                a.ShowDialog();
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {

                for (int i = 0; i < this.shipping_Fee1DataGridView.Rows.Count; i++)
                {
                    this.shipping_Fee1DataGridView.Rows[i].Cells[12].Value = true;
                }
            }
            else
            {
                for (int i = 0; i < this.shipping_Fee1DataGridView.Rows.Count; i++)
                {
                    this.shipping_Fee1DataGridView.Rows[i].Cells[12].Value = false;
                }
            }
         
        }

        private void button8_Click(object sender, EventArgs e)
        {
            UPDATESAP(textBox4.Text, textBox2.Text, textBox3.Text, comboBox4.SelectedValue.ToString());

            MessageBox.Show("更新成功");
            textBox4.Text = "";
            button2_Click(null, new EventArgs());
        }

        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = Get3();
        }





       
        

   
    }
}
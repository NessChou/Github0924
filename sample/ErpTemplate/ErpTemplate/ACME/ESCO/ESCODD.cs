using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.Web.UI;
using System.Net.Mime;

namespace ACME
{
    public partial class ESCODD : ACME.fmBase1
    {
        System.Net.Mail.Attachment data = null;
        private string FileName;
        public ESCODD()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            eSCO_DD1TableAdapter.Connection = MyConnection;
            eSCO_DD3TableAdapter.Connection = MyConnection;
            eSCO_DD5TableAdapter.Connection = MyConnection;
            eSCO_DD6TableAdapter.Connection = MyConnection;
            eSCO_DD7TableAdapter.Connection = MyConnection;
            eSCO_DD8TableAdapter.Connection = MyConnection;
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();

                eSCO.ESCO_DD1.RejectChanges();
      

            }
            catch
            {
            }
            return true;

        }
        public override void AfterCancelEdit()
        {
            WW();
        }
        public override void EndEdit()
        {
            WW();
        }
        public override void AfterEdit()
        {
            shippingCodeTextBox.ReadOnly = true;



        }
        public override void AfterAddNew()
        {
            WW();
        }
        public override void SetInit()
        {

            MyBS = eSCO_DD1BindingSource;
            MyTableName = "ESCO_DD1";
            MyIDFieldName = "ShippingCode";



        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "ES" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;


            string username = fmLogin.LoginID.ToString();


            uSERSTextBox.Text = username;
            this.eSCO_DD1BindingSource.EndEdit();
            kyes = null;


        }

        public override void FillData()
        {
            try
            {


                eSCO_DD1TableAdapter.Fill(eSCO.ESCO_DD1, MyID);
                eSCO_DD3TableAdapter.Fill(eSCO.ESCO_DD3, MyID);
                eSCO_DD5TableAdapter.Fill(eSCO.ESCO_DD5, MyID);
                eSCO_DD6TableAdapter.Fill(eSCO.ESCO_DD6, MyID);
                eSCO_DD7TableAdapter.Fill(eSCO.ESCO_DD7, MyID);
                eSCO_DD8TableAdapter.Fill(eSCO.ESCO_DD8, MyID);
                dataGridView1.DataSource = GETDD3(shippingCodeTextBox.Text);
            //    eSCO_DD2TableAdapter.Fill(eSCO.ESCO_DD2, MyID);
                //               string PATH = @"\\acmesrv01\SAP_Share\進流抽水站.jpg";
                //if (eSCO.ESCO_DD2.Rows.Count > 0)
                //{
                //    eSCO_DD2DataGridView[2,1].Value = Image.FromFile(PATH); 
                //}


                //DataGridViewImageColumn img = new DataGridViewImageColumn();

                //Image image = Image.FromFile(PATH);
                //img.Image = image;
                //eSCO_DD2DataGridView.Columns.Add(img);
                //img.HeaderText = "Image";
                //img.Name = "img";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static System.Data.DataTable GETDD3(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT LINE STEP,P1 站名  FROM ESCO_DD3 WHERE shippingcode=@shippingcode  ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " inv1 ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" inv1 "];
        }
        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {

                //shipping_OQUTDownloadBindingSource.MoveFirst();

                //for (int i = 0; i <= shipping_OQUTDownloadBindingSource.Count - 1; i++)
                //{
                //    DataRowView row1 = (DataRowView)shipping_OQUTDownloadBindingSource.Current;

                //    row1["seq"] = i;

                //    shipping_OQUTDownloadBindingSource.EndEdit();

                //    shipping_OQUTDownloadBindingSource.MoveNext();
                //}

                //for (int i = 0; i <= shipping_OQUTDownload2BindingSource.Count - 1; i++)
                //{
                //    DataRowView row1 = (DataRowView)shipping_OQUTDownload2BindingSource.Current;

                //    row1["seq"] = i;

                //    shipping_OQUTDownload2BindingSource.EndEdit();

                //    shipping_OQUTDownload2BindingSource.MoveNext();
                //}

    
                Validate();

                eSCO_DD3BindingSource1.MoveFirst();

                for (int i = 1; i <= eSCO_DD3BindingSource1.Count; i++)
                {
                    DataRowView row = (DataRowView)eSCO_DD3BindingSource1.Current;

                    row["LINE2"] = i;



                    eSCO_DD3BindingSource1.EndEdit();

                    eSCO_DD3BindingSource1.MoveNext();
                }

                eSCO_DD5BindingSource.MoveFirst();

                for (int i = 1; i <= eSCO_DD5BindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)eSCO_DD5BindingSource.Current;

                    row["LINE2"] = i;

                    eSCO_DD5BindingSource.EndEdit();

                    eSCO_DD5BindingSource.MoveNext();
                }

                eSCO_DD6BindingSource.MoveFirst();

                for (int i = 1; i <= eSCO_DD6BindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)eSCO_DD6BindingSource.Current;

                    row["LINE2"] = i;

                    eSCO_DD6BindingSource.EndEdit();

                    eSCO_DD6BindingSource.MoveNext();
                }

                eSCO_DD7BindingSource.MoveFirst();

                for (int i = 1; i <= eSCO_DD7BindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)eSCO_DD7BindingSource.Current;

                    row["LINE"] = i;

                    eSCO_DD7BindingSource.EndEdit();

                    eSCO_DD7BindingSource.MoveNext();
                }

                eSCO_DD8BindingSource.MoveFirst();

                for (int i = 1; i <= eSCO_DD8BindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)eSCO_DD8BindingSource.Current;

                    row["LINE"] = i;

                    eSCO_DD8BindingSource.EndEdit();

                    eSCO_DD8BindingSource.MoveNext();
                }


                eSCO_DD1TableAdapter.Connection.Open();


                eSCO_DD1BindingSource.EndEdit();
     
                tx = eSCO_DD1TableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(eSCO_DD1TableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;
   

                SqlDataAdapter Adapter2 = util.GetAdapter(eSCO_DD3TableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;


                SqlDataAdapter Adapter3 = util.GetAdapter(eSCO_DD5TableAdapter);
                Adapter3.UpdateCommand.Transaction = tx;
                Adapter3.InsertCommand.Transaction = tx;
                Adapter3.DeleteCommand.Transaction = tx;


                SqlDataAdapter Adapter4 = util.GetAdapter(eSCO_DD6TableAdapter);
                Adapter4.UpdateCommand.Transaction = tx;
                Adapter4.InsertCommand.Transaction = tx;
                Adapter4.DeleteCommand.Transaction = tx;


                SqlDataAdapter Adapter5 = util.GetAdapter(eSCO_DD7TableAdapter);
                Adapter5.UpdateCommand.Transaction = tx;
                Adapter5.InsertCommand.Transaction = tx;
                Adapter5.DeleteCommand.Transaction = tx;


                SqlDataAdapter Adapter6 = util.GetAdapter(eSCO_DD8TableAdapter);
                Adapter6.UpdateCommand.Transaction = tx;
                Adapter6.InsertCommand.Transaction = tx;
                Adapter6.DeleteCommand.Transaction = tx;

                eSCO_DD1TableAdapter.Update(eSCO.ESCO_DD1);
                eSCO.ESCO_DD1.AcceptChanges();

                eSCO_DD3TableAdapter.Update(eSCO.ESCO_DD3);
                eSCO.ESCO_DD3.AcceptChanges();

                eSCO_DD5TableAdapter.Update(eSCO.ESCO_DD5);
                eSCO.ESCO_DD5.AcceptChanges();

                eSCO_DD6TableAdapter.Update(eSCO.ESCO_DD6);
                eSCO.ESCO_DD6.AcceptChanges();

                eSCO_DD7TableAdapter.Update(eSCO.ESCO_DD7);
                eSCO.ESCO_DD7.AcceptChanges();

                eSCO_DD8TableAdapter.Update(eSCO.ESCO_DD8);
                eSCO.ESCO_DD8.AcceptChanges();

                tx.Commit();

                this.MyID = this.shippingCodeTextBox.Text;

                UpdateData = true;


            }
            catch (Exception ex)
            {
                if (tx != null)
                {

                    tx.Rollback();

                }


                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;

            }
            finally
            {
                this.eSCO_DD1TableAdapter.Connection.Close();

            }
            return UpdateData;
        }
   
        private void WW()
        {
            shippingCodeTextBox.ReadOnly = true;
            button2.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            checkBox1.Enabled = true;
            p20TextBox.ReadOnly = true;
        }

        //private void eSCO_DD2DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        //{

        //    int iRecs;

        //    iRecs = eSCO_DD2DataGridView.Rows.Count;
        //    e.Row.Cells["LINE"].Value = iRecs.ToString();
        //}

        private void ESCODD_Load(object sender, EventArgs e)
        {
            WW();

            DataGridViewLinkColumn column = new DataGridViewLinkColumn();
            column.Name = "Link";
            column.UseColumnTextForLinkValue = true;
            column.Text = "讀取檔案";
            column.LinkBehavior = LinkBehavior.HoverUnderline;
            column.TrackVisitedState = true;
            eSCO_DD7DataGridView.Columns.Add(column);

        }

        //private void eSCO_DD2DataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        //{

        //    try
        //    {
        //        if (!eSCO_DD2DataGridView.Rows[e.RowIndex].IsNewRow)
        //        {
        //            if (eSCO_DD2DataGridView.Columns[e.ColumnIndex].Name == "IMAGES")
        //            {
        //                string STEP = eSCO_DD2DataGridView.Rows[e.RowIndex].Cells["STEP"].Value.ToString();
        //                if (!String.IsNullOrEmpty(STEP))
        //                {
        //                    e.Value = Image.FromFile(@"\\acmesrv01\SAP_Share\ESCO\" + STEP + ".jpg");
        //                }

        //            }
        //        }
        //    }
        //    catch
        //    { 
            
        //    }
        //}

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FileName = openFileDialog1.FileName;


             //   WriteExcelAP2(FileName);
                WriteExcelAP3(FileName);
            }
        }
        public void AddAP2(string P1, string P2, string P3, string P4, string P5, string P6, string P7, string P8, string P9, string P10, string P11, string P12, string P13, string P14, string P15, string P16, string P17, string P18, string P19, string P20, string P21, string P22, int LINE, int LINE2)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ESCO_DD3(ShippingCode,P1,P2,P3,P4,P5,P6,P7,P8,P9,P10,P11,P12,P13,P14,P15,P16,P17,P18,P19,P20,P21,P22,LINE,LINE2) values(@ShippingCode,@P1,@P2,@P3,@P4,@P5,@P6,@P7,@P8,@P9,@P10,@P11,@P12,@P13,@P14,@P15,@P16,@P17,@P18,@P19,@P20,@P21,@P22,@LINE,@LINE2)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@P1", P1));
            command.Parameters.Add(new SqlParameter("@P2", P2));
            command.Parameters.Add(new SqlParameter("@P3", P3));
            command.Parameters.Add(new SqlParameter("@P4", P4));
            command.Parameters.Add(new SqlParameter("@P5", P5));
            command.Parameters.Add(new SqlParameter("@P6", P6));
            command.Parameters.Add(new SqlParameter("@P7", P7));
            command.Parameters.Add(new SqlParameter("@P8", P8));
            command.Parameters.Add(new SqlParameter("@P9", P9));
            command.Parameters.Add(new SqlParameter("@P10", P10));
            command.Parameters.Add(new SqlParameter("@P11", P11));
            command.Parameters.Add(new SqlParameter("@P12", P12));
            command.Parameters.Add(new SqlParameter("@P13", P13));
            command.Parameters.Add(new SqlParameter("@P14", P14));

            command.Parameters.Add(new SqlParameter("@P15", P15));
            command.Parameters.Add(new SqlParameter("@P16", P16));
            command.Parameters.Add(new SqlParameter("@P17", P17));
            command.Parameters.Add(new SqlParameter("@P18", P18));
            command.Parameters.Add(new SqlParameter("@P19", P19));
            command.Parameters.Add(new SqlParameter("@P20", P20));
            command.Parameters.Add(new SqlParameter("@P21", P21));
            command.Parameters.Add(new SqlParameter("@P22", P22));
            command.Parameters.Add(new SqlParameter("@LINE", LINE));
            command.Parameters.Add(new SqlParameter("@LINE2", LINE2));

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

        public void EE(string DD)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into RMA_PARAMS(PARAM_KIND,PARAM_NO,PARAM_DESC) values(@PARAM_KIND,@PARAM_NO,@PARAM_DESC)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PARAM_KIND", "ESCOT"));
            command.Parameters.Add(new SqlParameter("@PARAM_NO", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PARAM_DESC", DD));



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

        public void EE2()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE RMA_PARAMS WHERE PARAM_KIND='ESCOT'", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PARAM_NO", shippingCodeTextBox.Text));
    


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
        public void AddAP9(string PMONTH, string D1, string D2, string D3)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ESCO_DD9(PMONTH,D1,D2,D3) values(@PMONTH,@D1,@D2,@D3)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PMONTH", PMONTH));
            command.Parameters.Add(new SqlParameter("@D1", D1));
            command.Parameters.Add(new SqlParameter("@D2", D2));
            command.Parameters.Add(new SqlParameter("@D3", D3));


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
        public void AddAP10(string D1, string D2, string D3)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ESCO_DD10(D1,D2,D3) values(@D1,@D2,@D3)", connection);
            command.CommandType = CommandType.Text;
            
            command.Parameters.Add(new SqlParameter("@D1", D1));
            command.Parameters.Add(new SqlParameter("@D2", D2));
            command.Parameters.Add(new SqlParameter("@D3", D3));


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
        public void DEL9()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE ESCO_DD9 TRUNCATE TABLE ESCO_DD10", connection);
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
        public void DEL10()
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("TRUNCATE TABLE ESCO_DD10", connection);
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
        public void AddAP6(string PRICE, string QTY, string AMT, string MEMO, string ITEMNAME)
        {

            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into ESCO_DD6(ShippingCode,PRICE,QTY,AMT,MEMO,ITEMNAME) values(@ShippingCode,@PRICE,@QTY,@AMT,@MEMO,@ITEMNAME)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@AMT", AMT));
            command.Parameters.Add(new SqlParameter("@MEMO", MEMO));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
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
        private void WriteExcelAP2(string ExcelFile)
        {
            //  AddAP
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}




            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string P1;
                string DP1 = "";
                string P2;
                string P3;
                string P4;
                string P5;
                string P6;
                string P7;
                string P8;
                string P9;
                string P10;
                string P11;
                string P12;
                string P13;
                string P14;
                string P15;
                string P16;
                string P17;
                string P18;
                string P19;
                string P20;
                string P21;
                string P22 = "";

                int K1 = 0;
                int K2 = 0;
                for (int iRecord = 3; iRecord <= iRowCnt; iRecord++)
                {

                    K2++;
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    P1 = range.Text.ToString().Trim();


                    if (P1 == "")
                    {
                        P1 = DP1;

                    }
                    else
                    {
                        K1++;
                    }

                    System.Data.DataTable G1 = GeMAIN(P1);
                    if (G1.Rows.Count > 0)
                    {
                        P22 = G1.Rows[0][0].ToString();

                    }
                    else
                    {
                        P22 = P1;
                    }

                    int F1 = P22.IndexOf("MBR");
                    if (F1 != -1)
                    {
                        P22 = "MBR";
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    P2 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    P3 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 6]);
                    range.Select();
                    P4 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    P5 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    P6 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    P7 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    P8 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    P9 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 12]);
                    range.Select();
                    P10 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 13]);
                    range.Select();
                    P11 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 14]);
                    range.Select();
                    P12 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 15]);
                    range.Select();
                    P13 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 16]);
                    range.Select();
                    P14 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 17]);
                    range.Select();
                    P15 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 18]);
                    range.Select();
                    P16 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 19]);
                    range.Select();
                    P17 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 20]);
                    range.Select();
                    P18 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 21]);
                    range.Select();
                    P19 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 22]);
                    range.Select();
                    P20 = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 23]);
                    range.Select();
                    P21 = range.Text.ToString().Trim();
                    if (P1 != "")
                    {
                        DP1 = P1;
                    }

                    if (!String.IsNullOrEmpty(P2))
                    {
                        AddAP2(P1, P2, P3, P4, P5, P6, P7, P8, P9, P10, P11, P12, P13, P14, P15, P16, P17, P18, P19, P20, P21, P22, K1, K2);
                    }
                }




            }
            finally
            {

                //     string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
                //DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


                //     try
                //     {
                //         excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                //     }
                //     catch
                //     {
                //     }
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


                //  System.Diagnostics.Process.Start(NewFileName);


            }



        }
        private void WriteExcelAP3(string ExcelFile)
        {
            //  AddAP
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

            //if (!checkBox5.Checked)
            //{
            //    if (iRowCnt > 1000)
            //    {
            //        iRowCnt = 1000;
            //    }

            //}




            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {
                string ITEMNAME;

                string PRICE;
                string QTY;
                string AMT;
                string MEMO;
           
                for (int iRecord = 3; iRecord <= iRowCnt; iRecord++)
                {

              
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 1]);
                    range.Select();
                    ITEMNAME = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    PRICE = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    QTY = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    AMT = range.Text.ToString().Trim().Replace(",", "");

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    MEMO = range.Text.ToString().Trim();


                    if (!String.IsNullOrEmpty(ITEMNAME))
                    {
                        AddAP6(PRICE, QTY, AMT, MEMO, ITEMNAME);
                    }
                }




            }
            finally
            {

                //     string NewFileName = Path.GetDirectoryName(FileName) + "\\" +
                //DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);


                //     try
                //     {
                //         excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                //     }
                //     catch
                //     {
                //     }
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


                //  System.Diagnostics.Process.Start(NewFileName);


            }



        }
        private System.Data.DataTable GeMAIN(string DREF)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DMAIN   FROM ESCO_DD4 WHERE DREF=@DREF");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DREF", DREF));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        //private void eSCO_DD2DataGridView_MouseClick(object sender, MouseEventArgs e)
        //{
        //    eSCO_DD3TableAdapter.Fill(eSCO.ESCO_DD3, MyID);

        //    //try
        //    //{

        //    //    string da1 = sATT1DataGridView.SelectedRows[0].Cells["Seqno"].Value.ToString();
        //    //    for (int i = 0; i <= sATT2DataGridView.Rows.Count - 1; i++)
        //    //    {

        //    //        DataGridViewRow row;

        //    //        row = sATT2DataGridView.Rows[i];
        //    //        string a0 = row.Cells["dataGridViewTextBoxColumn3"].Value.ToString();

        //    //        if (da1 == a0)
        //    //        {
        //    //            sATT2DataGridView.FirstDisplayedScrollingRowIndex = i;
        //    //            break;
        //    //        }

        //    //    }
        //    //}
        //    //catch
        //    //{

        //    //}
        //}

        private void eSCO_DD3DataGridView_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void eSCO_DD3DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void 明細插入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = eSCO.ESCO_DD3;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;


            try
            {

                dt2.Rows.InsertAt(newCustomersRow, eSCO_DD3DataGridView.CurrentRow.Index);

                for (int j = 0; j <= eSCO_DD3DataGridView.Rows.Count - 2; j++)
                {
                    eSCO_DD3DataGridView.Rows[j].Cells[0].Value = (j + 1).ToString();
                }

                //this.eSCO_DD3BindingSource1.EndEdit();
                //this.eSCO_DD3TableAdapter.Update(eSCO.ESCO_DD3);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void 複製列ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt2 = eSCO.ESCO_DD3;
            DataRow newCustomersRow = dt2.NewRow();
            int i = eSCO_DD3DataGridView.CurrentRow.Index;

            DataRow drw = dt2.Rows[i];
            string sa = drw["shippingcode"].ToString();
            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["LINE2"] = 100;
            newCustomersRow["P1"] = drw["P1"];
            newCustomersRow["P2"] = drw["P2"];
            newCustomersRow["P3"] = drw["P3"];
            newCustomersRow["P4"] = drw["P4"];
            newCustomersRow["P5"] = drw["P5"];
            newCustomersRow["P6"] = drw["P6"];
            newCustomersRow["P7"] = drw["P7"];
            newCustomersRow["P8"] = drw["P8"];
            newCustomersRow["P9"] = drw["P9"];
            newCustomersRow["P10"] = drw["P10"];
            newCustomersRow["P11"] = drw["P11"];
            newCustomersRow["P12"] = drw["P12"];

            newCustomersRow["P13"] = drw["P13"];
            newCustomersRow["P14"] = drw["P14"];
            newCustomersRow["P15"] = drw["P15"];

            newCustomersRow["P16"] = drw["P16"];
            newCustomersRow["P17"] = drw["P17"];
            newCustomersRow["P18"] = drw["P18"];

            newCustomersRow["P19"] = drw["P19"];
            newCustomersRow["P20"] = drw["P20"];
            newCustomersRow["P21"] = drw["P21"];
            newCustomersRow["P22"] = drw["P22"];
            newCustomersRow["LINE"] = drw["LINE"];
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, eSCO_DD3DataGridView.CurrentRow.Index);

                for (int j = 0; j <= eSCO_DD3DataGridView.Rows.Count - 2; j++)
                {
                    eSCO_DD3DataGridView.Rows[j].Cells[0].Value = (j + 1).ToString();
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        public static System.Data.DataTable GETDDR(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT P1 系統名稱,P2 設備名稱,P3 設備編號,P4 設置數量,P5 運轉數量,P6 電壓規格,P7 設備銘牌功率,P8 設備銘牌馬力");
            sb.Append(" ,P9 設備運轉功率,P10 功率因素,P11 三相運轉電流,P12 三相動力線,P13 VFD,P14 銘牌揚程,P15 實際需求,P16 銘牌流量");
            sb.Append(" ,P17 需求流量,P18 閥門開度,P19 運轉條件,P20 感測器回控,P21 回控機制說明 FROM ESCO_DD3 WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public static System.Data.DataTable GETDDR2(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT P1 站名,P2 設備名稱,P3 設置數量,P4 運轉數量,P5 規劃運轉數量,P6 設備銘牌馬力,P7 設備運轉功率");
            sb.Append(" ,P8 設置總馬力,P9 總運轉功率,P10 運轉條件天,P11 運轉條件年,P12+'%' 節電率,P13 節能設施,P14 節電效益 FROM ESCO_DD5  WHERE SHIPPINGCODE=@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public static System.Data.DataTable GETDDR3(string SHIPPINGCODE, string DOCTYPE1)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DOCTYPE2+'-'+DOCTYPE1 類別,ITEMNAME 項次,PRICE 單價,QTY 數量,AMT 總價,MEMO 備註 FROM ESCO_DD6 WHERE SHIPPINGCODE=@SHIPPINGCODE AND DOCTYPE1=@DOCTYPE1");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@DOCTYPE1", DOCTYPE1));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public static System.Data.DataTable GETDDR4(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT CAST(CAST(P10 AS decimal)/CAST(100 AS DECIMAL) AS decimal(10,2)) P10,CAST(CAST(P12 AS decimal(10,4))/CAST(100 AS DECIMAL) AS decimal(10,4)) P12,* FROM ESCO_DD1   where shippingcode=@shippingcode  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDDR5(string shippingcode, string DOCTYPE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUM(CAST(AMT AS INT)) AMT FROM ESCO_DD6 WHERE DOCTYPE1=@DOCTYPE  AND shippingcode=@shippingcode  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDDR6(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT SUM(CAST(P14 AS INT)) FROM ESCO_DD5  WHERE shippingcode=@shippingcode  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        public static System.Data.DataTable GETDDR8(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT PYEAR-1911 PYEAR2,PYEAR,PMONTH   FROM ESCO_DD8 where shippingcode=@shippingcode ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDDR10()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT * FROM ESCO_DD9 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
         
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDDR10V1(string PMONTH)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT D1,D2,D3,CAST(D1 AS INT)+CAST(D2 AS INT)+CAST(D3 AS INT) D4 FROM ESCO_DD9 WHERE PMONTH=@PMONTH ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@PMONTH", PMONTH));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDDR12(string D1, string D2, string D3)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT D4 FROM ESCO_DD11 WHERE D1=@D1 AND D2=@D2 AND D3=@D3  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@D1", D1));
            command.Parameters.Add(new SqlParameter("@D2", D2));
            command.Parameters.Add(new SqlParameter("@D3", D3));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDDR13(string Y, string M)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT SUM(D1) D1,SUM(D2) D2,SUM(D3) D3 FROM Y_2004 WHERE YEAR(DATE_TIME)=@Y AND  MONTH(DATE_TIME)=@M  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@Y", Y));
            command.Parameters.Add(new SqlParameter("@M", M));
          
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDDR9(string shippingcode, string PYEAR, string PMONTH)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("    SELECT *  FROM ESCO_DD8 where shippingcode=@shippingcode AND PYEAR=@PYEAR AND PMONTH =@PMONTH ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@PYEAR", PYEAR));
            command.Parameters.Add(new SqlParameter("@PMONTH", PMONTH));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDD4(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("   SELECT * FROM ESCO_DD3  where shippingcode=@shippingcode ORDER BY LINE2 ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }
        public static System.Data.DataTable GETDDR7(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT CAST(CAST(P10 AS decimal)/CAST(100 AS DECIMAL) AS decimal(10,2)) P10,CAST(CAST(P12 AS decimal(10,4))/CAST(100 AS DECIMAL) AS decimal(10,4)) P12,* FROM ESCO_DD1   where shippingcode=@shippingcode  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.CommandTimeout = 0;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

        private void button3_Click(object sender, EventArgs e)
        {

            System.Data.DataTable dt1 = GETDD4(shippingCodeTextBox.Text);
            System.Data.DataTable dt2 = eSCO.ESCO_DD5;
            if (dt1.Rows.Count == 0)
            {
                MessageBox.Show("來源無資料，請先存檔");

                tabControl1.SelectedIndex = 0;

            }
            int h = 0;
        
            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                DataRow drw2 = dt2.NewRow();
                drw2["ShippingCode"] = shippingCodeTextBox.Text;
                drw2["LINE2"] = drw["LINE2"];
                //進流站_抽水泵_P101_124HP
                string P2 = drw["P1"].ToString() + "_" + drw["P2"].ToString() + "_" + drw["P3"].ToString().Replace("-", "") + "_" + drw["P8"].ToString() + "HP";
                drw2["P1"] = drw["P1"];
                drw2["P2"] = P2;
                drw2["P3"] = drw["P4"];
                drw2["P4"] = drw["P5"];
                drw2["P6"] = drw["P8"];
                drw2["P8"] = drw["P8"];
                dt2.Rows.Add(drw2);
            }





            eSCO_DD1BindingSource.EndEdit();
            eSCO_DD1TableAdapter.Update(eSCO.ESCO_DD1);
       eSCO.ESCO_DD1.AcceptChanges();

            eSCO_DD5BindingSource.EndEdit();
            eSCO_DD5TableAdapter.Update(eSCO.ESCO_DD5);
            eSCO.ESCO_DD5.AcceptChanges();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            System.Data.DataTable F1 = GETDDR(shippingCodeTextBox.Text);
            System.Data.DataTable F2 = GETDDR2(shippingCodeTextBox.Text);
            System.Data.DataTable F3 = GETDDR3(shippingCodeTextBox.Text,"設備費用");
            System.Data.DataTable F4 = GETDDR3(shippingCodeTextBox.Text, "安裝費用");
            System.Data.DataTable F5 = GETDDR4(shippingCodeTextBox.Text);
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ESCO\\DD.xlsx";
            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //產生 Excel ReportdataGridView1
            FESCODD(F1, F2, F3, F4, F5,ExcelTemplate, OutPutFile);
        }

        public void FESCODD(System.Data.DataTable OrderData, System.Data.DataTable OrderData2, System.Data.DataTable OrderData3, System.Data.DataTable OrderData4, System.Data.DataTable OrderData5, string ExcelFile, string OutPutFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            object SelectCell = null;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            //Microsoft.Office.Interop.Excel.Range range1 = null;



            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;
            Microsoft.Office.Interop.Excel.Range range = null;


            //Microsoft.Office.Interop.Excel.Range FixedRange = null;


            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {



                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(OrderData, sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            break;
                        }

                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != OrderData.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(OrderData, aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }


                Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
                excelSheet2.Activate();

                int iRowCnt2 = excelSheet2.UsedRange.Cells.Rows.Count;
                int iColCnt2 = excelSheet2.UsedRange.Cells.Columns.Count;



                string sTemp2 = string.Empty;
                string FieldValue2 = string.Empty;
                bool IsDetail2 = false;
                int DetailRow2 = 0;

                for (int iRecord = 1; iRecord <= iRowCnt2; iRecord++)
                {



                    for (int iField = 1; iField <= iColCnt2; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[iRecord, iField]);
                        range.Select();
                        sTemp2 = (string)range.Text;
                        sTemp2 = sTemp2.Trim();

                        if (CheckSerial(OrderData2, sTemp2, ref FieldValue2))
                        {
                            range.Value2 = FieldValue2;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp2))
                        {
                            IsDetail2 = true;
                            DetailRow2 = iRecord;
                            break;
                        }

                    }

                }

                if (DetailRow2 != 0)
                {

                    for (int aRow = 0; aRow <= OrderData2.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != OrderData2.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[DetailRow2, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt2; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[DetailRow2, iField]);
                            range.Select();
                            sTemp2 = (string)range.Text;
                            sTemp2 = sTemp2.Trim();

                            FieldValue2 = "";
                            SetRow(OrderData2, aRow, sTemp2, ref FieldValue2);

                            range.Value2 = FieldValue2;


                        }

                        DetailRow2++;
                    }

                }




                Microsoft.Office.Interop.Excel.Worksheet excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(3);
                excelSheet3.Activate();

                int iRowCnt3 = excelSheet3.UsedRange.Cells.Rows.Count;
                int iColCnt3 = excelSheet3.UsedRange.Cells.Columns.Count;



                string sTemp3 = string.Empty;
                string FieldValue3 = string.Empty;
                bool IsDetail3 = false;
                int DetailRow3 = 0;


                //5566

                for (int aRow = 0; aRow <= OrderData3.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != OrderData3.Rows.Count - 1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3 + 3, 1]);
                        range.EntireRow.Copy(oMissing);

                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt3; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3 + 3, iField]);
                        range.Select();
                        sTemp3 = (string)range.Text;
                        sTemp3 = sTemp3.Trim();

                        FieldValue3 = "";
                        SetRow(OrderData3, aRow, sTemp3, ref FieldValue3);

                        range.Value2 = FieldValue3;


                    }

                    DetailRow3++;
                }



                for (int aRow = 0; aRow <= OrderData4.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != OrderData4.Rows.Count - 1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3 + 5, 1]);
                        range.EntireRow.Copy(oMissing);

                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt3; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[DetailRow3 + 5, iField]);
                        range.Select();
                        sTemp3 = (string)range.Text;
                        sTemp3 = sTemp3.Trim();

                        FieldValue3 = "";
                        SetRow(OrderData4, aRow, sTemp3, ref FieldValue3);

                        range.Value2 = FieldValue3;


                    }

                    DetailRow3++;
                }


                Microsoft.Office.Interop.Excel.Worksheet excelSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(4);
                excelSheet4.Activate();

                int iRowCnt4 = excelSheet4.UsedRange.Cells.Rows.Count;
                int iColCnt4 = excelSheet4.UsedRange.Cells.Columns.Count;



                string sTemp4 = string.Empty;
                string FieldValue4 = string.Empty;

                for (int iRecord = 1; iRecord <= iRowCnt4; iRecord++)
                {



                    for (int iField = 1; iField <= iColCnt4; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet4.UsedRange.Cells[iRecord, iField]);
                        range.Select();
                        sTemp4 = (string)range.Text;
                        sTemp4 = sTemp4.Trim();

                        if (CheckSerial(OrderData5, sTemp4, ref FieldValue4))
                        {
                            range.Value2 = FieldValue4;
                        }



                    }

                }
          
            }
            finally
            {


                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
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



                System.Diagnostics.Process.Start(OutPutFile);

            }

        }
        public void FESCODD2(System.Data.DataTable OrderData,string ExcelFile, string OutPutFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            object SelectCell = null;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            //Microsoft.Office.Interop.Excel.Range range1 = null;



            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;
            Microsoft.Office.Interop.Excel.Range range = null;


            //Microsoft.Office.Interop.Excel.Range FixedRange = null;


            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;




             
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {



                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(OrderData, sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }

            

                    }

                }

         
                System.Data.DataTable G1 = GETDDR5(shippingCodeTextBox.Text, "設備費用");
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                range.Select();
                range.Value2 = G1.Rows[0][0].ToString();

                System.Data.DataTable G2 = GETDDR5(shippingCodeTextBox.Text, "安裝費用");
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[6, 3]);
                range.Select();
                range.Value2 = G2.Rows[0][0].ToString();

                double  E1 = Convert.ToDouble(G1.Rows[0][0]);
                double E2 = Convert.ToDouble(G2.Rows[0][0]);
                double E3 = Convert.ToDouble(p12TextBox.Text);


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[9, 3]);
                range.Select();
                range.Value2 = (((E1 + E2) + ((E1 + E2) * 0.06)) * E3).ToString();

                System.Data.DataTable G3 = GETDDR6(shippingCodeTextBox.Text);
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[31, 1]);
                range.Select();
                range.Value2 = G3.Rows[0][0].ToString();
            }
            finally
            {


                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
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



                System.Diagnostics.Process.Start(OutPutFile);

            }

        }
        public void FESCODD5(System.Data.DataTable OrderData, string ExcelFile, string OutPutFile, string FLAG)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            object SelectCell = null;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            //Microsoft.Office.Interop.Excel.Range range1 = null;



            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);



            // progressBar1.Maximum = iRowCnt;
            Microsoft.Office.Interop.Excel.Range range = null;


            //Microsoft.Office.Interop.Excel.Range FixedRange = null;

            Microsoft.Office.Interop.Excel.Worksheet excelSheet2 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(4);
            excelSheet2.Activate();
            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;

       
                DEL9();
                System.Data.DataTable LL = GETDDR8(shippingCodeTextBox.Text);
                for (int N = 0; N <= 11; N++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[4 + N, 8]);
                    range.Select();
                    string R1 = LL.Rows[N]["PMONTH"].ToString();
                    string PYEAR = LL.Rows[N]["PYEAR"].ToString();
                    System.Data.DataTable S1 = GETDDR13(PYEAR, R1);
                    string D1 = S1.Rows[0]["D1"].ToString();
                    string D2 = S1.Rows[0]["D2"].ToString();
                    string D3 = S1.Rows[0]["D3"].ToString();
                    range.Value2 = R1;

                    string d1 = "";
                    string d2 = "";
                    string e1 = "";
                    string e2 = "";
                    string f1 = "";
                    string f2 = "";
                    if (R1 == "6" || R1 == "7" || R1 == "8" || R1 == "9")
                    {
                        d1 = D1;
                        d2 = "0";
                        e1 = D2;
                        e2 = "0";
                        f1 = D3;
                        f2 = "0";
                    }
                    else
                    {
                        d1 = "0";
                        d2 = D1;
                        e1 = "0";
                        e2 = D2;
                        f1 = "0";
                        f2 = D3;
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[4 + N, 9]);
                    range.Select();
                    //     string d1 = range.Value.ToString();
                    range.Value2 = d1;
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[4 + N, 10]);
                    range.Select();
                    range.Value2 = e1;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[4 + N, 11]);
                    range.Select();
                    range.Value2 = f1;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[4 + N, 12]);
                    range.Select();
                    range.Value2 = d2;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[4 + N, 13]);
                    range.Select();
                    range.Value2 = e2;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet2.UsedRange.Cells[4 + N, 14]);
                    range.Select();
                    range.Value2 = f2;

                    AddAP9(R1, (Convert.ToInt16(d1) + Convert.ToInt16(d2)).ToString(), (Convert.ToInt16(e1) + Convert.ToInt16(e2)).ToString(), (Convert.ToInt16(f1) + Convert.ToInt16(f2)).ToString());
                }


                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
                excelSheet.Activate();
                System.Data.DataTable L1 = GETDDR8(shippingCodeTextBox.Text);
                for (int N = 1; N <= 12; N++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[12, N + 2]);
                    range.Select();
                    string YEAR = L1.Rows[N - 1]["PYEAR2"].ToString();
                    string YEAR2 = L1.Rows[N - 1]["PYEAR"].ToString();
                    range.Value2 = YEAR;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[13, N + 2]);
                    range.Select();
                    string MONTH = L1.Rows[N - 1]["PMONTH"].ToString().Trim();
                  
                    range.Value2 = util.EXDD(MONTH);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[14, N + 2]);
                    range.Select();
                    range.Value2 = util.EXDD(MONTH);
                    System.Data.DataTable L2 = GETDDR9(shippingCodeTextBox.Text, YEAR2, MONTH);
                    if (MONTH.Length == 1)

                    {
                        MONTH = "0" + MONTH;
                    }
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[15, N + 2]);
                    range.Select();
                    range.Value2 = YEAR2 + "/" + MONTH + "/1";
                    string YM = YEAR2 + MONTH;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[16, N + 2]);
                    range.Select();
                    range.Value2 = GetMenu.DLast3(YM);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[24, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P1"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[25, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P2"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[26, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P3"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[27, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P4"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[28, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P5"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[29, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P6"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[30, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P7"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[31, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P8"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[33, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P9"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[34, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P10"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[35, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P11"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[36, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P12"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[38, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P13"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[39, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P14"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[40, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P15"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[41, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P16"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[43, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P17"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[44, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P18"].ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[55, N + 2]);
                    range.Select();
                    range.Value2 = L2.Rows[0]["P19"].ToString();
                   

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[51, N + 2]);
                    range.Select();
                    string Q1=range.Value.ToString();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[53, N + 2]);
                    range.Select();
                    string Q2= range.Value.ToString();

                    int n;

                       if (int.TryParse(Q1, out n) && int.TryParse(Q2, out n) && int.TryParse(L2.Rows[0]["P19"].ToString(), out n))
                       {
                           range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[57, N + 2]);
                           range.Select();
                           range.Value2 = (Convert.ToDecimal(Convert.ToInt32(Q1) + Convert.ToInt32(Q2)) / Convert.ToDecimal(L2.Rows[0]["P19"].ToString())).ToString();
                       }
                    string M2 = L1.Rows[N - 1]["PMONTH"].ToString().Trim();
                    System.Data.DataTable LL2S = GETDDR10V1(M2);
                    string S1, S2, S3, S4, S5, S6 = "";
                    string D1 = LL2S.Rows[0]["D1"].ToString();
                    string D2 = LL2S.Rows[0]["D2"].ToString();
                    string D3 = LL2S.Rows[0]["D3"].ToString();
                    if (M2 == "6" || M2 == "7" || M2 == "8" || M2 == "9")
                    {
                        S1 = D1;
                        S2 = D2;
                        S3 = D3;
                        S4 = "";
                        S5 = "";
                        S6 = "";
                    }
                    else
                    {
                        S1 = ""; 
                        S2 = ""; 
                        S3 = ""; 
                        S4 = D1;
                        S5 = D2;
                        S6 = D3;
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[18, N + 2]);
                    range.Select();
                    range.Value2 = S1;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[19, N + 2]);
                    range.Select();
                    range.Value2 = S2;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[20, N + 2]);
                    range.Select();
                    range.Value2 = S3;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[21, N + 2]);
                    range.Select();
                    range.Value2 = S4;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[22, N + 2]);
                    range.Select();
                    range.Value2 = S5;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[23, N + 2]);
                    range.Select();
                    range.Value2 = S6;
                }


                Microsoft.Office.Interop.Excel.Worksheet excelSheet5 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(2);
                excelSheet5.Activate();

                for (int N = 1; N <= 12; N++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet5.UsedRange.Cells[2, N + 2]);
                    range.Select();
                    string YEAR = L1.Rows[N - 1]["PYEAR2"].ToString();
                    range.Value2 = YEAR;

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet5.UsedRange.Cells[3, N + 2]);
                    range.Select();
                    string MONTH = L1.Rows[N - 1]["PMONTH"].ToString().Trim();

                    range.Value2 = util.EXDD(MONTH);

                }

                               Microsoft.Office.Interop.Excel.Worksheet excelSheet3 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(3);
                               excelSheet3.Activate();
                         
                               System.Data.DataTable LL2 = GETDDR10();
                               for (int N = 0; N <= 11; N++)
                               {
                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[15, 3 + N]);
                                   range.Select();
                                   range.Value2 = LL2.Rows[N]["D1"].ToString();

                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[16, 3 + N]);
                                   range.Select();
                                   range.Value2 = LL2.Rows[N]["D2"].ToString();

                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[17, 3 + N]);
                                   range.Select();
                                   range.Value2 = LL2.Rows[N]["D3"].ToString();



                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[2, N + 3]);
                                       range.Select();
                                       string YEAR = L1.Rows[N]["PYEAR2"].ToString();
                                       string YEAR2 = L1.Rows[N]["PYEAR"].ToString();
                                       range.Value2 = YEAR;

                                       range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[3, N + 3]);
                                       range.Select();
                                       string MONTH = L1.Rows[N]["PMONTH"].ToString().Trim();

                                       range.Value2 = util.EXDD(MONTH);

                                       int 週間 = Convert.ToInt16(LL2.Rows[N]["D1"]);

                                       int 週六 = Convert.ToInt16(LL2.Rows[N]["D2"]);
          
                                       int 假日 = Convert.ToInt16(LL2.Rows[N]["D3"]);
                                       string DTYPE = "非夏月";
                                       if (MONTH == "6" || MONTH == "7" || MONTH == "8" || MONTH == "9")
                                       {
                                           DTYPE = "夏月";
                                       }

                                       int S1 = Convert.ToInt16(GETDDR12("尖峰", DTYPE, "週間").Rows[0][0]);
                                       int S2 = Convert.ToInt16(GETDDR12("尖峰", DTYPE, "週六").Rows[0][0]);
                                       int S3 = Convert.ToInt16(GETDDR12("尖峰", DTYPE, "假日").Rows[0][0]);

                                       int S4 = Convert.ToInt16(GETDDR12("半尖峰", DTYPE, "週間").Rows[0][0]);
                                       int S5 = Convert.ToInt16(GETDDR12("半尖峰", DTYPE, "週六").Rows[0][0]);
                                       int S6 = Convert.ToInt16(GETDDR12("半尖峰", DTYPE, "假日").Rows[0][0]);

                                       int S7 = Convert.ToInt16(GETDDR12("週六半尖", DTYPE, "週間").Rows[0][0]);
                                       int S8 = Convert.ToInt16(GETDDR12("週六半尖", DTYPE, "週六").Rows[0][0]);
                                       int S9 = Convert.ToInt16(GETDDR12("週六半尖", DTYPE, "假日").Rows[0][0]);

                                       int S10 = Convert.ToInt16(GETDDR12("離峰", DTYPE, "週間").Rows[0][0]);
                                       int S11 = Convert.ToInt16(GETDDR12("離峰", DTYPE, "週六").Rows[0][0]);
                                       int S12 = Convert.ToInt16(GETDDR12("離峰", DTYPE, "假日").Rows[0][0]);
                                       //尖峰
                                       range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[18, N + 3]);
                                       range.Select();
                                       range.Value2 = (週間 * S1 + 週六 * S2 + 假日 * S3).ToString();
                                       //半尖峰
                                       range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[19, N + 3]);
                                       range.Select();
                                       range.Value2 = (週間 * S4 + 週六 * S5 + 假日 * S6).ToString();
                                       //週六半尖
                                       range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[20, N + 3]);
                                       range.Select();
                                       range.Value2 = (週間 * S7 + 週六 * S8 + 假日 * S9).ToString();
                                       //離峰
                                       range = ((Microsoft.Office.Interop.Excel.Range)excelSheet3.UsedRange.Cells[21, N + 3]);
                                       range.Select();
                                       range.Value2 = (週間 * S10 + 週六 * S11 + 假日 * S12).ToString();

                                   
                               }

                               int G1 = Convert.ToInt16(p17TextBox.Text);
                              
                               int G3 = Convert.ToInt16(p20TextBox.Text);
                               int G4 = G3 - 12;
                               Microsoft.Office.Interop.Excel.Worksheet excelSheet4 = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(6);
                            
                               excelSheet4.Activate();


                               if (G4 > 0)
                               {
                                   for (int N = 0; N <= G4 - 1; N++)
                                   {
                                       range = ((Microsoft.Office.Interop.Excel.Range)excelSheet4.UsedRange.Cells[59, 1]);
                                       range.EntireRow.Copy(oMissing);

                                       range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                           oMissing);
                                   }
                               }

                               for (int N = 0; N <= G3-1; N++)
                               {
                                   int G2 = Convert.ToInt16(p19TextBox.Text);
                                   if (N == 0)
                                   {
                                       G2 = 0;
                                   }
                          

                                   G1 = G1 - G2;
                                

                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet4.UsedRange.Cells[30, 2]);
                                   range.Select();
                                   range.Value2 = G1.ToString();

                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet4.UsedRange.Cells[51, 14]);
                                   range.Select();
                                   string F1 = range.Value.ToString();

                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet4.UsedRange.Cells[52, 14]);
                                   range.Select();
                                   string F2 = range.Value.ToString();



                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet4.UsedRange.Cells[58+N, 2]);
                                   range.Select();
                                   range.Value2 = G1.ToString();

                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet4.UsedRange.Cells[58+N, 3]);
                                   range.Select();
                                   range.Value2 = F1;

                                   range = ((Microsoft.Office.Interop.Excel.Range)excelSheet4.UsedRange.Cells[58+N, 4]);
                                   range.Select();
                                   range.Value2 = F2;

                                   //range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[58 + N, 1]);
                                   //range.EntireRow.Copy(oMissing);

                                   //range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                   //    oMissing);
                               }

                            

                      
            


            }
            finally
            {


                try
                {
                    excelSheet2.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
                //Quit
                excelApp.Quit();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet2);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                range = null;
                excelApp = null;
                excelBook = null;
                excelSheet2 = null;

                System.GC.Collect();
                //可以將 Excel.exe 清除
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("產生一個檔案->" + NewFileName);


                if (FLAG == "Y")
                {
                    System.Diagnostics.Process.Start(OutPutFile);
                }

            }

        }
        private string GetExePath()
        {
            return Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
        }

        public void FESCODD3(System.Data.DataTable OrderData,string ExcelFile, string OutPutFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            object SelectCell = null;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            //Microsoft.Office.Interop.Excel.Range range1 = null;



            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;
            Microsoft.Office.Interop.Excel.Range range = null;


            //Microsoft.Office.Interop.Excel.Range FixedRange = null;


            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {



                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(OrderData, sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            break;
                        }

                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData.Rows.Count - 1; aRow++)
                    {

                        //最後一筆不作
                        if (aRow != OrderData.Rows.Count - 1)
                        {

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, 1]);
                            range.EntireRow.Copy(oMissing);

                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                                oMissing);
                        }


                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow, iField]);
                            range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(OrderData, aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }



            }
            finally
            {


                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
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



                System.Diagnostics.Process.Start(OutPutFile);

            }

        }

        public void FESCODD4(System.Data.DataTable OrderData3, System.Data.DataTable OrderData4,  string ExcelFile, string OutPutFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            excelApp.DisplayAlerts = false;
            excelApp.Visible = false;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            object SelectCell = null;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            //Microsoft.Office.Interop.Excel.Range range1 = null;



            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


            //取得 Excel 的使用區域


            // progressBar1.Maximum = iRowCnt;
            Microsoft.Office.Interop.Excel.Range range = null;


            //Microsoft.Office.Interop.Excel.Range FixedRange = null;


            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Activate();

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;

  




                int iRowCnt3 = excelSheet.UsedRange.Cells.Rows.Count;
                int iColCnt3 = excelSheet.UsedRange.Cells.Columns.Count;



                string sTemp3 = string.Empty;
                string FieldValue3 = string.Empty;
                bool IsDetail3 = false;
                int DetailRow3 = 0;


                //5566

                for (int aRow = 0; aRow <= OrderData3.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != OrderData3.Rows.Count - 1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow3 + 3, 1]);
                        range.EntireRow.Copy(oMissing);

                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt3; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow3 + 3, iField]);
                        range.Select();
                        sTemp3 = (string)range.Text;
                        sTemp3 = sTemp3.Trim();

                        FieldValue3 = "";
                        SetRow(OrderData3, aRow, sTemp3, ref FieldValue3);

                        range.Value2 = FieldValue3;


                    }

                    DetailRow3++;
                }



                for (int aRow = 0; aRow <= OrderData4.Rows.Count - 1; aRow++)
                {

                    //最後一筆不作
                    if (aRow != OrderData4.Rows.Count - 1)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow3 + 5, 1]);
                        range.EntireRow.Copy(oMissing);

                        range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                            oMissing);
                    }


                    for (int iField = 1; iField <= iColCnt3; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow3 + 5, iField]);
                        range.Select();
                        sTemp3 = (string)range.Text;
                        sTemp3 = sTemp3.Trim();

                        FieldValue3 = "";
                        SetRow(OrderData4, aRow, sTemp3, ref FieldValue3);

                        range.Value2 = FieldValue3;


                    }

                    DetailRow3++;
                }



            }
            finally
            {


                try
                {
                    excelSheet.SaveAs(OutPutFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //增加一個 Close
                excelBook.Close(oMissing, oMissing, oMissing);
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



                System.Diagnostics.Process.Start(OutPutFile);

            }

        }
        public static void SetRow(System.Data.DataTable OrderData, int iRow, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return;
            }
            if (sData.Substring(0, 2) == "[[")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[iRow][FieldName]);
            }

        }

        public static bool IsDetailRow(string sData)
        {

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "[[")
            {

                return true;
            }
            //}
            return false;
        }
        public static bool CheckSerial(System.Data.DataTable OrderData, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "<<")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master 固定第一筆
                FieldValue = Convert.ToString(OrderData.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }

        private void eSCO_DD5DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (eSCO_DD5DataGridView.Columns[e.ColumnIndex].Name == "設備銘牌馬力" ||
             eSCO_DD5DataGridView.Columns[e.ColumnIndex].Name == "設置數量")
                {
                    decimal Q1 = 0;
                    decimal Q2 = 0;

                    Q1 = Convert.ToDecimal(this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["設備銘牌馬力"].Value);
                    Q2 = Convert.ToDecimal(this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["設置數量"].Value);
                    this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["設置總馬力"].Value = (Q1 * Q2).ToString();

                }
                if (eSCO_DD5DataGridView.Columns[e.ColumnIndex].Name == "設備運轉功率" ||
      eSCO_DD5DataGridView.Columns[e.ColumnIndex].Name == "規劃運轉數量")
                {
                    decimal  Q1 = 0;
                    decimal Q2 = 0;

                    Q1 = Convert.ToDecimal(this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["設備運轉功率"].Value);
                    Q2 = Convert.ToDecimal(this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["規劃運轉數量"].Value);
                    this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["總運轉功率"].Value = (Q1 * Q2).ToString();

                }

                if (eSCO_DD5DataGridView.Columns[e.ColumnIndex].Name == "總運轉功率" ||
eSCO_DD5DataGridView.Columns[e.ColumnIndex].Name == "運轉條件天" ||
                    eSCO_DD5DataGridView.Columns[e.ColumnIndex].Name == "運轉條件年" ||
                    eSCO_DD5DataGridView.Columns[e.ColumnIndex].Name == "節電率")
                {
                    decimal Q1 = 0;
                    decimal Q2 = 0;
                    decimal Q3 = 0;
                    decimal Q4 = 0;
                    Q1 = Convert.ToDecimal(this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["總運轉功率"].Value);
                    Q2 = Convert.ToDecimal(this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["運轉條件天"].Value);
                    Q3 = Convert.ToDecimal(this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["運轉條件年"].Value);
                    Q4 = Convert.ToDecimal(this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["節電率"].Value) / 100;
                    this.eSCO_DD5DataGridView.Rows[e.RowIndex].Cells["節電效益"].Value = Convert.ToInt32(Q1 * Q2 * Q3 * Q4).ToString();

                }
            }
            catch
            {

            }
         //   if (shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "Quantity" ||
         //shipping_ItemDataGridView.Columns[e.ColumnIndex].Name == "ItemPrice")
         //   {
         //       decimal iQuantity = 0;
         //       decimal iUnitPrice = 0;

         //       iQuantity = Convert.ToInt32(this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["Quantity"].Value);
         //       iUnitPrice = Convert.ToDecimal(this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["ItemPrice"].Value);
         //       this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["ItemAmount"].Value = (iQuantity * iUnitPrice).ToString();

         //   }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            DELETEFILE();
 
            string NAME = "";
            if (tabControl1.SelectedIndex == 0)
            {
                System.Data.DataTable F1 = GETDDR(shippingCodeTextBox.Text);
                FileName = lsAppDir + "\\Excel\\ESCO\\DD1.xlsx";

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                FESCODD3(F1, FileName, OutPutFile);

            }
            if (tabControl1.SelectedIndex == 1)
            {
                System.Data.DataTable F2 = GETDDR2(shippingCodeTextBox.Text);
                FileName = lsAppDir + "\\Excel\\ESCO\\DD2.xlsx";

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                FESCODD3(F2, FileName, OutPutFile);
            }
            if (tabControl1.SelectedIndex == 2)
            {

                System.Data.DataTable F3 = GETDDR3(shippingCodeTextBox.Text, "設備費用");
                System.Data.DataTable F4 = GETDDR3(shippingCodeTextBox.Text, "安裝費用");
                FileName = lsAppDir + "\\Excel\\ESCO\\DD3.xlsx";

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                FESCODD4(F3, F4, FileName, OutPutFile);
            }
            if (tabControl1.SelectedIndex == 3)
            {

                System.Data.DataTable F5 = GETDDR4(shippingCodeTextBox.Text);
                FileName = lsAppDir + "\\Excel\\ESCO\\DD4.xlsx";

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                FESCODD2(F5, FileName, OutPutFile);
            }
            if (tabControl1.SelectedIndex == 4)
            {

                System.Data.DataTable F5 = GETDDR4(shippingCodeTextBox.Text);
                string DOCNUM = dOCNAMETextBox.Text;
                int K1 = DOCNUM.IndexOf("-");
                if (K1 != -1)
                {
                    DOCNUM = DOCNUM.Substring(0, K1);
                }
                FileName = lsAppDir + "\\Excel\\ESCO\\" + DOCNUM + ".xlsx";

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);
                string FLAG = "Y";
                if (checkBox1.Checked)
                {
                    FLAG = "N";
                }
                FESCODD5(F5, FileName, OutPutFile,FLAG);
                if (checkBox1.Checked)
                {
                    MAIL();
                }
            }
        }
        private void MAIL()
        {

            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\MailTemplates\\LC.html";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();
            objReader.Dispose();

            StringWriter writer = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);

            string user=fmLogin.LoginID.ToString();
            template = template.Replace("##F1##", "Hi " + user +",");
            template = template.Replace("##F2##", "請看附件");
            template = template.Replace("##AA##", "");
            MailMessage message = new MailMessage();


            message.To.Add(fmLogin.LoginID.ToString() + "@ACMEPOINT.COM");




            message.Subject = "電力需量分析試算";
            message.Body = template;
            string OutPutFile = lsAppDir + "\\Excel\\temp";
            string[] filenames = Directory.GetFiles(OutPutFile);
            foreach (string file in filenames)
            {

                string m_File = "";

                m_File = file;
                data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);

                //附件资料
                ContentDisposition disposition = data.ContentDisposition;


                // 加入邮件附件
                message.Attachments.Add(data);


            }


            message.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.Send(message);
            data.Dispose();
            message.Attachments.Dispose();


            MessageBox.Show("寄信成功");
                    
        }

        private void MAIL2()
        {

            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\MailTemplates\\LC.html";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();
            objReader.Dispose();

            StringWriter writer = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);

            string user = fmLogin.LoginID.ToString();
            template = template.Replace("##F1##", "Hi " + user + ",");
            template = template.Replace("##F2##", "請看附件");
            template = template.Replace("##AA##", "");
            MailMessage message = new MailMessage();


            message.To.Add(fmLogin.LoginID.ToString() + "@ACMEPOINT.COM");




            message.Subject = "施工計劃書";
            message.Body = template;
            string OutPutFile = lsAppDir + "\\Excel\\temp";
            string[] filenames = Directory.GetFiles(OutPutFile);
        //    MessageBox.Show(OutPutFile.ToString());
            foreach (string file in filenames)
            {
              //  MessageBox.Show(file.ToString());
                int SS1 = file.IndexOf("docx");
                if (SS1 != -1)
                {
                    string m_File = "";

                    m_File = file;
                    data = new System.Net.Mail.Attachment(m_File, MediaTypeNames.Application.Octet);

                    //附件资料
                    ContentDisposition disposition = data.ContentDisposition;


                    // 加入邮件附件
                    message.Attachments.Add(data);
                }


            }


            message.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.Send(message);
            data.Dispose();
            message.Attachments.Dispose();


            MessageBox.Show("寄信成功");

        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
  
        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                string server = "//acmesrv01//SAP_Share//LC//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);

                if (result == DialogResult.OK)
                {


                    string file = opdf.FileName;
                    bool FF1 = getrma.UploadFile(file, server, false);
                    if (FF1 == false)
                    {
                        return;
                    }
                    System.Data.DataTable dt1 = eSCO.ESCO_DD7;

                    DataRow drw = dt1.NewRow();
                    string DOCTYPE = "";
                    if (dataGridView1.SelectedRows.Count > 0)
                    {
                        DataGridViewRow row;
                        StringBuilder sb = new StringBuilder();

                        row = dataGridView1.SelectedRows[0];

                        DOCTYPE = row.Cells["站名"].Value.ToString();

                    }
                    else
                    {
                        MessageBox.Show("請選擇站名");
                        return;
                    }

                    drw["DOCTYPE"] = DOCTYPE;
                    drw["shippingcode"] = shippingCodeTextBox.Text;
                    drw["LINE"] = (eSCO_DD7DataGridView.Rows.Count).ToString();
                    drw["filename"] = filename;
                    drw["path"] = @"\\acmesrv01\SAP_Share\LC\" + filename;
                    dt1.Rows.Add(drw);
                    this.eSCO_DD7BindingSource.EndEdit();
                    this.eSCO_DD7TableAdapter.Update(eSCO.ESCO_DD7);
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void eSCO_DD7DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;


                if (dgv.Columns[e.ColumnIndex].Name == "Link")
                {
                    System.Data.DataTable dt1 = eSCO.ESCO_DD7;
                    int i = e.RowIndex;
                    DataRow drw = dt1.Rows[i];

                    string aa = drw["path"].ToString();


                    System.Diagnostics.Process.Start(aa);

                    DataGridViewLinkCell cell =

                        (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                    cell.LinkVisited = true;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public DataTable GetPATH()
        {

            SqlConnection connection = globals.Connection;

            string sql = "SELECT PARAM_NO  FROM RMA_PARAMS WHERE PARAM_KIND='COPYPATH3'";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "right");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["right"];
        }
        private void button5_Click(object sender, EventArgs e)
        {
            string DD="";
            if (checkBox1.Checked)
            {
                DD = "Y";
            }
            EE2();
            EE(DD);
            DELETEFILE();
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            string NewFileName = lsAppDir + "\\WindowsFormsApp2.exe";
            System.Diagnostics.Process.Start(NewFileName);
            if (checkBox1.Checked)
            {
                string OutPutFile = lsAppDir + "\\Excel\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
               MessageBox.Show("寄信中");
                foreach (string file in filenames)
                {
                   // MessageBox.Show(file.ToString());
                 


                }

                MAIL2();
            }
    
    
        }

        private System.Data.DataTable GETS1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT P1,P2,P3,P4,P5,P7,P8,P10  FROM ESCO_DD5  ");

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

        private void eSCO_DD8DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = eSCO_DD8DataGridView.Rows.Count;
            e.Row.Cells["LINE8"].Value = iRecs.ToString();
        }

        private void p19TextBox_TextChanged(object sender, EventArgs e)
        {
            SA();
        }

        private void SA()
        {
            try
            {
                int S1 = Convert.ToInt16(p17TextBox.Text);
                int S2 = Convert.ToInt16(p18TextBox.Text);
                int S3 = Convert.ToInt16(p19TextBox.Text);

                p20TextBox.Text = (((S1 - S2) / S3) + 1).ToString();
            }
            catch { }
        }

        private void p18TextBox_TextChanged(object sender, EventArgs e)
        {
            SA();
        }

        private void p17TextBox_TextChanged(object sender, EventArgs e)
        {
            SA();
        }
    }
}

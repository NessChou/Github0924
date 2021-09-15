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
using System.Net.Mail;
using System.Reflection;
using System.Web.UI;
using System.Collections;
using System.Net.Mime;
using System.Runtime.InteropServices;

namespace ACME
{
    public partial class SHICAR : ACME.fmBase1
    {
        public string PublicString;
        public SHICAR()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;

            shipping_CARTableAdapter.Connection = MyConnection;
            shipping_CAR2TableAdapter.Connection = MyConnection;
            shipping_CAR3TableAdapter.Connection = MyConnection;
            shipping_CAR4TableAdapter.Connection = MyConnection;
            shipping_CARDownloadTableAdapter.Connection = MyConnection;
        }
        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
            comboBox1.SelectedValue = -1;
            comboBox2.SelectedValue = -1;
            comboBox3.SelectedValue = -1;
            comboBox4.SelectedValue = -1;
          
            cLOSETIMETextBox.Text = "";
            cLOSESTATUSTextBox.Text = "";
           // cLOSEDCheckBox.Checked = false;
            cLOSETYPEComboBox.SelectedValue = -1;
            comboBox5.SelectedValue = -1;
        }

        public override void EndEdit()
        {
            WW();
        }
    
        public override void STOP()
        {

            if (shipping_CAR2DataGridView.Rows.Count > 1)
            {
                for (int j = 0; j <= shipping_CAR2DataGridView.Rows.Count - 2; j++)
                {

                    string DOCDATE = shipping_CAR2DataGridView.Rows[j].Cells["DOCDATE2"].Value.ToString();

                    if (String.IsNullOrEmpty(DOCDATE))
                    {
                        MessageBox.Show("請輸入預交日期");
                        this.SSTOPID = "1";
                        return;
                    }

                }

            }
        }

        public override void AfterEndEdit()
        {
            CalcTotals2();
            SHIPC();
     
            if (cLOSETYPEComboBox.Text == "工單結案" && cLOSEDCheckBox.Checked)
            {
                DialogResult result;
                result = MessageBox.Show("是否要更新工單結案?", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    CalcTotals();
                    MessageBox.Show("結案資料已更新");
                }
            }

        }
        private void WW()
        {

            shippingCodeTextBox.ReadOnly = true ;
        }
        public override void AfterCancelEdit()
        {
            WW();
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                ship.Shipping_CAR.RejectChanges();
                ship.Shipping_CAR2.RejectChanges();
                ship.Shipping_CAR3.RejectChanges();
                ship.Shipping_CAR4.RejectChanges();
                ship.Shipping_CARDownload.RejectChanges();
            }
            catch
            {
            }
            return true;

        }
        public override void SetInit()
        {

            MyBS = shipping_CARBindingSource;
            MyTableName = "Shipping_CAR";
            MyIDFieldName = "ShippingCode";
        }
        public override void AfterAddNew()
        {
            WW();

        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "SC" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;

            System.Data.DataTable J1 = GETOHEM(fmLogin.LoginID.ToString().Trim());
            if (J1.Rows.Count > 0)
            {
                createNameTextBox.Text = J1.Rows[0][0].ToString();

            }
            cLOSEDCheckBox.Checked = false;
            this.shipping_CARBindingSource.EndEdit();
            kyes = null;
            cLOSETYPEComboBox.Text = "併單結案";
        }
        public override void FillData()
        {
            try
            {
                if (!String.IsNullOrEmpty(PublicString))
                {
                    MyID = PublicString.Trim();
                }

                CalcTotals2();
                shipping_CARTableAdapter.Fill(ship.Shipping_CAR, MyID);
                shipping_CAR2TableAdapter.Fill(ship.Shipping_CAR2, MyID);
                shipping_CAR3TableAdapter.Fill(ship.Shipping_CAR3, MyID);
                shipping_CAR4TableAdapter.Fill(ship.Shipping_CAR4, MyID);
                shipping_CARDownloadTableAdapter.Fill(ship.Shipping_CARDownload, MyID);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void CalcTotals()
        {
            int i = this.shipping_CAR2DataGridView.Rows.Count - 2;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                string JOBNO = shipping_CAR2DataGridView.Rows[iRecs].Cells["JOBNO2"].Value.ToString();

      
                UPDATECLOSE(JOBNO, cLOSETIMETextBox.Text);

            }

        }

        private void CalcTotals2()
        {
            int i = this.shipping_CAR2DataGridView.Rows.Count - 2;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                string JOBNO = shipping_CAR2DataGridView.Rows[iRecs].Cells["JOBNO2"].Value.ToString();

                string 費用 = "";
                string 進度 = "";

                System.Data.DataTable DT1 = GETMEMO(JOBNO);
                if (DT1.Rows.Count > 0)
                {
                    費用 = DT1.Rows[0]["費用"].ToString();
                    進度 = DT1.Rows[0]["進度"].ToString();
                    int T1 = 費用.IndexOf("費用紀錄詳見");
                    int T2 = 進度.IndexOf("進度紀錄詳見");
                    if (T1 == -1)
                    {
                        if (String.IsNullOrEmpty(費用))
                        {
                            費用 ="費用紀錄詳見" + shippingCodeTextBox.Text;
                        }
                        else
                        {
                            費用 = 費用 + Environment.NewLine + "費用紀錄詳見" + shippingCodeTextBox.Text;
                        }
                        UPDATEMEMO1(JOBNO, 費用);
                    }

                    if (T2 == -1)
                    {
                        if (String.IsNullOrEmpty(進度))
                        {
                            進度 = "進度紀錄詳見" + shippingCodeTextBox.Text;
                        }
                        else
                        {
                            進度 = 進度 + Environment.NewLine + "進度紀錄詳見" + shippingCodeTextBox.Text;
                        }
                        UPDATEMEMO2(JOBNO, 進度);
                    }
                }

            }

        }
        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {


                Validate();

                shipping_CAR2BindingSource.MoveFirst();
                for (int i = 1; i <= shipping_CAR2BindingSource.Count; i++)
                {
                    DataRowView row2 = (DataRowView)shipping_CAR2BindingSource.Current;

                    row2["SeqNo"] = i;



                    shipping_CAR2BindingSource.EndEdit();

                    shipping_CAR2BindingSource.MoveNext();
                }



                shipping_CARTableAdapter.Connection.Open();

                shipping_CARBindingSource.EndEdit();
                shipping_CAR2BindingSource.EndEdit();
                shipping_CAR3BindingSource.EndEdit();
                shipping_CAR4BindingSource.EndEdit();
                shipping_CARDownloadBindingSource.EndEdit();

                tx = shipping_CARTableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(shipping_CARTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter1 = util.GetAdapter(shipping_CAR2TableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter3 = util.GetAdapter(shipping_CAR3TableAdapter);
                Adapter3.UpdateCommand.Transaction = tx;
                Adapter3.InsertCommand.Transaction = tx;
                Adapter3.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter4 = util.GetAdapter(shipping_CAR4TableAdapter);
                Adapter4.UpdateCommand.Transaction = tx;
                Adapter4.InsertCommand.Transaction = tx;
                Adapter4.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter5 = util.GetAdapter(shipping_CARDownloadTableAdapter);
                Adapter5.UpdateCommand.Transaction = tx;
                Adapter5.InsertCommand.Transaction = tx;
                Adapter5.DeleteCommand.Transaction = tx;


                shipping_CARTableAdapter.Update(ship.Shipping_CAR);
                ship.Shipping_CAR.AcceptChanges();

                shipping_CAR2TableAdapter.Update(ship.Shipping_CAR2);
                ship.Shipping_CAR2.AcceptChanges();

                shipping_CAR3TableAdapter.Update(ship.Shipping_CAR3);
                ship.Shipping_CAR3.AcceptChanges();

                shipping_CAR4TableAdapter.Update(ship.Shipping_CAR4);
                ship.Shipping_CAR4.AcceptChanges();

                shipping_CARDownloadTableAdapter.Update(ship.Shipping_CARDownload);
                ship.Shipping_CARDownload.AcceptChanges();

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
                this.shipping_CARTableAdapter.Connection.Close();

            }



            return UpdateData;
        }
        public override void SAVE()
        {
            shipping_CARBindingSource.EndEdit();
            shipping_CAR2BindingSource.EndEdit();
            shipping_CAR3BindingSource.EndEdit();
            shipping_CAR4BindingSource.EndEdit();
            shipping_CARDownloadBindingSource.EndEdit();



            shipping_CARTableAdapter.Update(ship.Shipping_CAR);
            shipping_CAR2TableAdapter.Update(ship.Shipping_CAR2);
            shipping_CAR3TableAdapter.Update(ship.Shipping_CAR3);
            shipping_CAR4TableAdapter.Update(ship.Shipping_CAR4);
            shipping_CARDownloadTableAdapter.Update(ship.Shipping_CARDownload);


            ship.Shipping_CAR.AcceptChanges();
            ship.Shipping_CAR2.AcceptChanges();
            ship.Shipping_CAR3.AcceptChanges();
            ship.Shipping_CAR4.AcceptChanges();
            ship.Shipping_CARDownload.AcceptChanges();



            MessageBox.Show("儲存成功");

        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text != "")
            {

                string aa = GetMenu.SPLITDOC(textBox1.Text.Replace("\r\n", ""));

                System.Data.DataTable dt1 = GetMenu.GetSHICAR(aa);
                System.Data.DataTable dt2 = ship.Shipping_CAR2;
                if (dt1.Rows.Count > 0)
                {
                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();
                        drw2["JOBNO"] = drw["SHIPPINGCODE"];
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["QTY"] = drw["QTY"];
                        drw2["NET"] = drw["NET"];
                        drw2["GROSS"] = drw["GROSS"];
                        drw2["PACKAGE"] = drw["PACKAGE"];
                        drw2["CARDNAME"] = drw["CARDNAME"];
                        drw2["OWNER"] = drw["OWNER"];
                        drw2["CBM"] = drw["CBM"];
                        string DOC = drw["DOC"].ToString();
                        string 類別 = drw["類別"].ToString();
                        string ADD1 = drw["ADD1"].ToString();
                        int FA = ADD1.IndexOf("正航");
                        if (FA == -1)
                        {
                            if (類別 == "銷售")
                            {
                                System.Data.DataTable T1 = GetMenu.GetSHICARSA(DOC);
                                if (T1.Rows.Count > 0)
                                {
                                    drw2["SA"] = T1.Rows[0][0].ToString();
                                }
                            }
                            if (類別 == "採購")
                            {
                                System.Data.DataTable T1 = GetMenu.GetSHICARSA2(DOC);
                                if (T1.Rows.Count > 0)
                                {
                                    drw2["SA"] = T1.Rows[0][0].ToString();
                                }
                            }
                        }
                        dt2.Rows.Add(drw2);
                    }

                    textBox1.Text = "";
                }
                else
                {
                    MessageBox.Show("沒有資料");
                }
            }
        }
        public void SHIPC()
        {

            System.Data.DataTable dt2 = GetMenu.GetSHICAR2(shippingCodeTextBox.Text);

            if (dt2.Rows.Count > 0)
            {
                DELETECAR31(shippingCodeTextBox.Text);
                for (int S = 0; S <= dt2.Rows.Count - 1; S++)
                {
                    DataRow drw = dt2.Rows[S];
                    string CM = drw["CM"].ToString();
                    string PACKAGE = drw["PACKAGE"].ToString();
                    string CM2 = drw["CM2"].ToString();
                    int T1 = CM.IndexOf("/");
                    int TG = CM2.IndexOf("@");
                    if (T1 != -1)
                    {
                        string[] arrurl = CM.Split(new Char[] { '/' });

                        foreach (string i in arrurl)
                        {
                            string PACK = "1";
                            string DD = i;
                            int T2 = DD.IndexOf("*");
                            if (T2 != -1)
                            {
                                PACK = DD.Substring(T2 + 1, DD.Length - T2 - 1);

                                DD = DD.Substring(0, T2);
                            }
                            INSCAR3(shippingCodeTextBox.Text, DD, PACK);
                        }

                    }
                    else if (TG != -1)
                    {
                        INSCAR3(shippingCodeTextBox.Text, CM, PACKAGE);
                    }
                    else
                    {
                        string PACK = "1";
                        int T2 = CM.IndexOf("*");
                        if (T2 != -1)
                        {
                            PACK = CM.Substring(T2 + 1, CM.Length - T2 - 1);
                            CM = CM.Substring(0, T2);
                        }

                        INSCAR3(shippingCodeTextBox.Text, CM, PACK);
                    }

                }


                System.Data.DataTable dt3 = GetMenu.GetSHICAR3(shippingCodeTextBox.Text);
                if (dt3.Rows.Count > 0)
                {
                    DELETECAR31(shippingCodeTextBox.Text);
                    for (int S = 0; S <= dt3.Rows.Count - 1; S++)
                    {

                        DataRow drw = dt3.Rows[S];
                        INSCAR3(shippingCodeTextBox.Text, drw["CM"].ToString(), drw["PACKAGE"].ToString());
                    }

                }
            }
            shipping_CAR3TableAdapter.Fill(ship.Shipping_CAR3, MyID);
        
        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = ship.Shipping_CAR2;
            if (dt2.Rows.Count > 0)
            {
               for  (int i = 0; i <= dt2.Rows.Count - 1; i++)
                {
                    string JOBNO = dt2.Rows[i]["JOBNO"].ToString();

                    System.Data.DataTable dt2T = GetMenu.GetSHICART(JOBNO);

                    if (dt2T.Rows.Count > 0)
                    {
                        DataRow drw = dt2T.Rows[0];
                        decimal NET = Convert.ToDecimal(drw["NET"]);
                        decimal GROSS = Convert.ToDecimal(drw["GROSS"]);
                        string PACKAGE = drw["PACKAGE"].ToString();
                        string CARDNAME = drw["CARDNAME"].ToString();
                        int QTY = Convert.ToInt32(drw["QTY"]);
                        string OWNER = drw["OWNER"].ToString();
                        string DOC = drw["DOC"].ToString();
                        string 類別 = drw["類別"].ToString();
                        string CBM = drw["CBM"].ToString();
                        string SA = "";
                        if (類別 == "銷售" && DOC.Length < 8)
                        {
                            System.Data.DataTable T1 = GetMenu.GetSHICARSA(DOC);
                            if (T1.Rows.Count > 0)
                            {
                                SA = T1.Rows[0][0].ToString();
                            }
                        }
                        if (類別 == "採購" && DOC.Length < 8)
                        {
                            System.Data.DataTable T1 = GetMenu.GetSHICARSA2(DOC);
                            if (T1.Rows.Count > 0)
                            {
                                SA = T1.Rows[0][0].ToString();
                            }
                        }
                        if (類別 == "銷售" && DOC.Length >= 8)
                        {
                            //CHOICE

                            System.Data.DataTable T1 = GetMenu.GetSHICARSACHOICE(DOC);
                            if (T1.Rows.Count > 0)
                            {
                                SA = T1.Rows[0][0].ToString();
                            }
                        }   

                        UPDATEPACK(NET, GROSS, PACKAGE, JOBNO, shippingCodeTextBox.Text, SA, OWNER, CARDNAME, QTY, CBM);

                       
                    
                    }
                }
            }
            shipping_CAR2TableAdapter.Fill(ship.Shipping_CAR2, MyID);
            SHIPC();

        }


        public void UPDATEPACK(decimal Net, decimal Gross, string Package, string JOBNO, string SHIPPINGCODE, string SA, string OWNER, string CardName, int QTY, string CBM)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" UPDATE  Shipping_CAR2 SET Net=@Net,Gross=@Gross,Package=@Package,SA=@SA,OWNER=@OWNER,CardName=@CardName,QTY=@QTY,CBM=@CBM  WHERE JOBNO=@JOBNO AND SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@Net", Net));
            command.Parameters.Add(new SqlParameter("@Gross", Gross));
            command.Parameters.Add(new SqlParameter("@Package", Package));
            command.Parameters.Add(new SqlParameter("@JOBNO", JOBNO));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@SA", SA));
            command.Parameters.Add(new SqlParameter("@OWNER", OWNER));
            command.Parameters.Add(new SqlParameter("@CardName", CardName));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@CBM", CBM));
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

        public void UPDATECLOSE(string SHIPPINGCODE, string buCardname)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" UPDATE  shipping_Main SET buCardcode='Checked',buCardname=@buCardname  where  SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@buCardname", buCardname));
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

        public void UPDATEMEMO1(string SHIPPINGCODE, string MEMO1)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" UPDATE  shipping_Main SET MEMO1=@MEMO1  where  SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@MEMO1", MEMO1));
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

        public void UPDATEMEMO2(string SHIPPINGCODE, string notifyMemo)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" UPDATE  shipping_Main SET notifyMemo=@notifyMemo  where  SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@notifyMemo", notifyMemo));
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
        public void INSCAR3(string ShippingCode, string MeasurmentCM, string Package)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into Shipping_CAR3(ShippingCode,MeasurmentCM,Package) values(@ShippingCode,@MeasurmentCM,@Package)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));
            command.Parameters.Add(new SqlParameter("@MeasurmentCM", MeasurmentCM));
            command.Parameters.Add(new SqlParameter("@Package", Package));


            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }
        public void DELETECAR21(string ShippingCode)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE Shipping_CAR2 WHERE ShippingCode=@ShippingCode ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));



            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }
        public void DELETECAR31(string ShippingCode)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE Shipping_CAR3 WHERE ShippingCode=@ShippingCode ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", ShippingCode));



            try
            {
                connection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                connection.Close();
            }
        }
        private void comboBox2_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt4 = GetMenu.Getwarehouse();

            comboBox2.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt4.Rows[i][1]));
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            wHFROMTextBox.Text = comboBox2.Text;
        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt4 = GetMenu.Getwarehouse();

            comboBox1.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt4.Rows[i][1]));
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            wHTOTextBox.Text = comboBox1.Text;
        }

        private void shipping_CAR2DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["Net2"].Value = 0;
            e.Row.Cells["QTY2"].Value = 0;
            e.Row.Cells["Gross2"].Value = 0;
            e.Row.Cells["Package2"].Value = "0";
        }

        private void comboBox3_MouseClick(object sender, MouseEventArgs e)
        {
            //WH
            System.Data.DataTable dt3 = GetOHEMSHIP1();


            comboBox3.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox3.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

            createNameTextBox.Text = comboBox3.Text;
        }

        private void shipping_CAR2DataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }

            //if (e.RowIndex >= gB_POTATODataGridView.Rows.Count - 1)
            //    return;

        }

        private void shipping_CAR3DataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }

        private void shipping_CAR4DataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            DataGridView dgv = (DataGridView)sender;

            using (SolidBrush b = new SolidBrush(dgv.RowHeadersDefaultCellStyle.ForeColor))
            {
                e.Graphics.DrawString((e.RowIndex + 1).ToString(), e.InheritedRowStyle.Font,
                    b, e.RowBounds.Location.X + 20, e.RowBounds.Location.Y + 6);
            }
        }
        private StringBuilder htmlMessageBody(DataGridView dg)
        {

            string KeyValue = "";

            string tmpKeyValue = "";

            StringBuilder strB = new StringBuilder();

            if (dg.Rows.Count == 0)
            {
                strB.AppendLine("<table class='GridBorder' cellspacing='0'");
                strB.AppendLine("<tr><td>***  查無資料  ***</td></tr>");
                strB.AppendLine("</table>");

                return strB;

            }

            //create html & table
            //strB.AppendLine("<html><body><center><table border='1' cellpadding='0' cellspacing='0'>");
            strB.AppendLine("<table class='GridBorder'  border='1' cellspacing='0' rules='all'  style='border-collapse:collapse;'>");
            strB.AppendLine("<tr class='HeaderBorder'>");
            //cteate table header
            for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
            {
                strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
            }
            strB.AppendLine("</tr>");

            //GridView 要設成不可加入及編輯．．不然會多一行空白
            for (int i = 0; i <= dg.Rows.Count - 1; i++)
            {

                if (KeyValue != dg.Rows[i].Cells[0].Value.ToString())
                {



                    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                    tmpKeyValue = KeyValue;
                }
                else
                {
                    tmpKeyValue = "";
                }


                if (i % 2 == 0)
                {
                    strB.AppendLine("<tr class='RowBorder'>");
                }
                else
                {
                    strB.AppendLine("<tr class='AltRowBorder'>");
                }



                // foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)
                DataGridViewCell dgvc;
                //foreach (DataGridViewCell dgvc in dg.Rows[i].Cells)

                if (string.IsNullOrEmpty(tmpKeyValue))
                {
                    strB.AppendLine("<td>&nbsp;</td>");
                }
                else
                {
                    strB.AppendLine("<td>" + tmpKeyValue + "</td>");
                }


                for (int d = 1; d <= dg.Rows[i].Cells.Count - 1; d++)
                {
                    dgvc = dg.Rows[i].Cells[d];
                    // HttpUtility.HtmlDecode("&nbsp;&nbsp;&nbsp;")

                    if (dgvc.ValueType == typeof(Int32))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Int32 x = Convert.ToInt32(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0") + "</td>");
                        }


                    }

                    else if (dgvc.ValueType == typeof(Decimal) || dgvc.ValueType == typeof(Double))
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {
                            Decimal x = Convert.ToDecimal(dgvc.Value);
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0.00") + "</td>");
                        }


                    }
                    else
                    {
                        //if (Convert.IsDBNull(dgvc.Value.ToString()))
                        if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                        {
                            // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                            strB.AppendLine("<td>&nbsp;</td>");
                        }
                        else
                        {

                            if (dg.Columns[dgvc.ColumnIndex].HeaderText.IndexOf("日期") >= 0)
                            {
                                if (dgvc.Value.ToString() == "0")
                                {
                                    strB.AppendLine("<td>&nbsp;</td>");
                                }
                                else
                                {

                                    string sDate = dgvc.Value.ToString().Substring(0, 4) + "/" +
                                                 dgvc.Value.ToString().Substring(4, 2) + "/" +
                                                 dgvc.Value.ToString().Substring(6, 2);


                                    strB.AppendLine("<td>" + sDate + "</td>");
                                }
                            }
                            else
                            {
                                strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                            }
                        }

                    }


                }
                strB.AppendLine("</tr>");

            }
  
            strB.AppendLine("</table>");
            return strB;
        }
        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("收件人地址為" + textBox2.Text + "是否要寄出", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {


                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                System.Data.DataTable A1 = GETMAIL1();
                System.Data.DataTable A2 = GETMAIL2();
                System.Data.DataTable A3 = GETMAIL3();
                dataGridView1.DataSource = A1;
                dataGridView2.DataSource = A2;
                dataGridView3.DataSource = A3;
                  string MailContent = htmlMessageBody(dataGridView1).ToString();
                  string MailContent2 = htmlMessageBody(dataGridView2).ToString();
                  string MailContent3 = htmlMessageBody(dataGridView3).ToString();
                FileName = lsAppDir + "\\MailTemplates\\併車.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();
                objReader.Dispose();

                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
                template = template.Replace("##SQUT1##", "併單");
                template = template.Replace("##SQUT2##", MailContent);
                template = template.Replace("##SQUT3##", "棧板尺寸");
                template = template.Replace("##SQUT4##", MailContent2);
                template = template.Replace("##SQUT5##", "車型尺寸");
                template = template.Replace("##SQUT6##", MailContent3);
                MailMessage message = new MailMessage();

                string aa = textBox2.Text;

                message.To.Add(new MailAddress(aa));

                message.Subject = "併車通知: 工單號碼:" + shippingCodeTextBox.Text;
                message.Body = template;

                //格式為 Html
                message.IsBodyHtml = true;

                SmtpClient client = new SmtpClient();
                try
                {
                    client.Send(message);

                    MessageBox.Show("寄信成功");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }


            }
        }

        private void SHICAR_Load(object sender, EventArgs e)
        {
            textBox2.Text = fmLogin.LoginID.ToString() + "@acmepoint.com";

            WW();
        }

        private void shipping_CAR4DataGridView_Paint(object sender, PaintEventArgs e)
        {
            System.Drawing.Rectangle r2 = this.shipping_CAR4DataGridView.GetCellDisplayRectangle(1, -1, true);

            System.Drawing.Rectangle r22 = this.shipping_CAR4DataGridView.GetCellDisplayRectangle(2, -1, true);

            System.Drawing.Rectangle r23 = this.shipping_CAR4DataGridView.GetCellDisplayRectangle(3, -1, true);

            //get the column header 

            r2.X += 1;

            r2.Y += 1;

            r2.Width = r2.Width + r22.Width + r23.Width - 2;

            r2.Height = r2.Height / 2 - 2;

            e.Graphics.FillRectangle(new SolidBrush(this.shipping_CAR4DataGridView.ColumnHeadersDefaultCellStyle.BackColor), r2);

            StringFormat format1 = new StringFormat();

            format1.Alignment = StringAlignment.Center;

            format1.LineAlignment = StringAlignment.Center;

            e.Graphics.DrawString("內徑尺寸(CM)", this.shipping_CAR4DataGridView.ColumnHeadersDefaultCellStyle.Font, new SolidBrush(this.shipping_CAR4DataGridView.ColumnHeadersDefaultCellStyle.ForeColor), r2, format1);

            Brush gridBrush = new SolidBrush(this.shipping_CAR4DataGridView.GridColor);
            Pen gridLinePen = new Pen(gridBrush);
            e.Graphics.DrawLine(gridLinePen, r2.X, r2.Y + r2.Height, r2.X + r2.Width, r2.Y + r2.Height);


        }

        private void shipping_CAR4DataGridView_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex == -1 && e.ColumnIndex > -1)
            {
                e.PaintBackground(e.CellBounds, false);
                System.Drawing.Rectangle r2 = e.CellBounds;
                r2.Y += e.CellBounds.Height / 2;
                r2.Height = e.CellBounds.Height / 2; e.PaintContent(r2);
                e.Handled = true;
            }
        }

        private void shipping_CAR4DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void shipping_CAR4DataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int scrollPosition = e.RowIndex;

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                DataGridViewColumn column = (sender as DataGridView).Columns[e.ColumnIndex];
                if (column.Name == "colEdit")
                {
                    shipping_CAR4BindingSource.EndEdit();
                    shipping_CAR4TableAdapter.Update(ship.Shipping_CAR4);

                    DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;
                    if (row != null)
                    {
                        string ID = Convert.ToString(row["ID"]);
                        SHICAROCRD form = new SHICAROCRD(ID);
                        if (form.ShowDialog() == DialogResult.OK)
                        {
                            shipping_CAR4TableAdapter.Fill(ship.Shipping_CAR4, MyID);
                            try
                            {
                                (sender as DataGridView).CurrentCell = (sender as DataGridView)[0, scrollPosition];
                            }
                            catch
                            {

                            }
                        }

                    }
                }

                if (column.Name == "colEdit2")
                {
                    shipping_CAR4BindingSource.EndEdit();
                    shipping_CAR4TableAdapter.Update(ship.Shipping_CAR4);

                    DataRowView row = (DataRowView)(sender as DataGridView).Rows[e.RowIndex].DataBoundItem;
                    if (row != null)
                    {
                        string ID = Convert.ToString(row["ID"]);
                        string CARTYPE = Convert.ToString(row["CARTYPE"]);
                        SHICAROCRD2 form = new SHICAROCRD2(ID, CARTYPE);
                        if (form.ShowDialog() == DialogResult.OK)
                        {
                            shipping_CAR4TableAdapter.Fill(ship.Shipping_CAR4, MyID);
                            try
                            {
                                (sender as DataGridView).CurrentCell = (sender as DataGridView)[0, scrollPosition];
                            }
                            catch
                            {

                            }
                        }

                    }
                }

            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

            cARTYPETextBox.Text = comboBox4.Text;
        }

        private void comboBox4_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt4 = GetMenu.GetBU("SHICARTYPE");

            comboBox4.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox4.Items.Add(Convert.ToString(dt4.Rows[i][1]));
            }
        }


        private void button9_Click(object sender, EventArgs e)
        {
            string DIR = "//acmesrv01//SAP_Share//shipping//";
            string PATH = @"\\acmesrv01\SAP_Share\shipping\";
            string f = "c";
            string[] filebType = Directory.GetDirectories(DIR);
            string dd = DateTime.Now.ToString("yyyyMM");
            string tt = DIR + dd;
            foreach (string fileaSize in filebType)
            {

                if (fileaSize == tt)
                {
                    f = "d";

                }

            }
            if (f == "c")
            {
                Directory.CreateDirectory(tt);
            }
            try
            {
                string server = DIR + dd + "//";
                

       
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    foreach (String fileS in openFileDialog1.FileNames)
                    {
                        string filename = Path.GetFileName(fileS);
                        string file = fileS;
                        bool FF1 = getrma.UploadFile(file, server, false);
                        if (FF1 == false)
                        {
                            MessageBox.Show("檔案沒有上傳");
                            return;
                        }
                        System.Data.DataTable dt1 = ship.Shipping_CARDownload;

                        DataRow drw = dt1.NewRow();
                        drw["ShippingCode"] = shippingCodeTextBox.Text;
                        drw["seq"] = (shipping_CARDownloadDataGridView.Rows.Count).ToString();
                        drw["filename"] = filename;
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
                        drw["path"] = PATH + de + filename;
                        dt1.Rows.Add(drw);

                        shipping_CARDownloadBindingSource.MoveFirst();

                        for (int i = 0; i <= shipping_CARDownloadBindingSource.Count - 1; i++)
                        {
                            DataRowView rowd = (DataRowView)shipping_CARDownloadBindingSource.Current;

                            rowd["seq"] = i + 1;

                            shipping_CARDownloadBindingSource.EndEdit();
                            shipping_CARDownloadBindingSource.MoveNext();
                        }

                        this.shipping_CARDownloadBindingSource.EndEdit();
                        this.shipping_CARDownloadTableAdapter.Update(ship.Shipping_CARDownload);
                        this.ship.Shipping_CARDownload.AcceptChanges();


                        MessageBox.Show("上傳成功");
                    }
                }



            }
            catch (Exception ex)
            {

            }
        }

        private void shipping_CARDownloadDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK2")
                {

                    System.Data.DataTable dt1 = ship.Shipping_CARDownload;
                    int i = e.RowIndex;
                    DataRow drw = dt1.Rows[i];
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                    string aa = drw["path"].ToString();
                    string filename = drw["filename"].ToString();
                    string NewFileName = lsAppDir + "\\EXCEL\\temp\\" + filename;

                    System.IO.File.Copy(aa, NewFileName, true);
                    System.Diagnostics.Process.Start(NewFileName);

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

        private void cLOSEDCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (cLOSEDCheckBox.Checked)
            {

                cLOSETIMETextBox.Text = DateTime.Now.ToString("yyyyMMdd");
                cLOSESTATUSTextBox.Text = "已結";
            }
            else
            {
                cLOSETIMETextBox.Text = "";
                cLOSESTATUSTextBox.Text = "未結";
            }
        }

        private void comboBox5_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("SHIPSTATUS");

            comboBox5.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox5.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            cLOSESTATUSTextBox.Text = comboBox5.Text;

            if (cLOSESTATUSTextBox.Text == "已結")
            {
                cLOSETIMETextBox.Text = GetMenu.Day();
                cLOSEDCheckBox.Checked = true;
            }
            else if (cLOSESTATUSTextBox.Text == "取消")
            {
                cLOSETIMETextBox.Text = GetMenu.Day();
                cLOSEDCheckBox.Checked = false;
            }
            else
            {
                cLOSETIMETextBox.Text = "";
                cLOSEDCheckBox.Checked = false;
            }
        }

        private void shipping_CAR2DataGridView_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (shipping_CAR2DataGridView.SelectedRows.Count > 0)
            {

                string da = shipping_CAR2DataGridView.SelectedRows[0].Cells["JOBNO2"].Value.ToString();

                fmShip a = new fmShip();
                a.PublicString = da;

                a.ShowDialog();
            }
        }

        private System.Data.DataTable GETMAIL1()
        {
            StringBuilder sb = new StringBuilder();


            SqlConnection MyConnection = globals.Connection;

            sb.Append(" SELECT JOBNO 工單號碼,CardName 客戶名稱,QTY 數量,Net,Gross,Package,CBM,[OWNER] 船務所有人,DOCDATE 預交日期 FROM shipping_CAR2 WHERE ShippingCode =@SHIPPINGCODE");
            sb.Append(" UNION ALL");
            sb.Append(" SELECT '' 工單號碼,'總計' 客戶名稱,SUM(CAST(QTY AS INT)) 數量,SUM(CAST(Net AS decimal(18,2))),SUM(CAST(Gross AS DECIMAL(18,2))),SUM(CAST(Package AS decimal(18,2))),SUM(CAST(CBM AS decimal(18,2))),'' 船務所有人,'' 預交日期 FROM shipping_CAR2 WHERE  ShippingCode =@SHIPPINGCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }
        private System.Data.DataTable GETMAIL2()
        {
            StringBuilder sb = new StringBuilder();


            SqlConnection MyConnection = globals.Connection;
            sb.Append(" SELECT MeasurmentCM,Package  FROM shipping_CAR3 WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }
        private System.Data.DataTable GETMAIL3()
        {
            StringBuilder sb = new StringBuilder();


            SqlConnection MyConnection = globals.Connection;
            sb.Append(" SELECT CARSIZE '車型(8Tor40)',CARSIZEL 長,CARSIZEW 寬,CARSIZEH 高,CARTYPE 廠商  FROM shipping_CAR4 WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }
        private System.Data.DataTable GETMEMO(string SHIPPINGCODE)
        {
            StringBuilder sb = new StringBuilder();


            SqlConnection MyConnection = globals.Connection;
            sb.Append(" SELECT MEMO1 費用,notifyMemo 進度  FROM SHIPPING_MAIN WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }
        private System.Data.DataTable GETOHEM(string HOMETEL)
        {
            StringBuilder sb = new StringBuilder();


            SqlConnection MyConnection = globals.shipConnection;
            sb.Append(" SELECT CASE HOMETEL WHEN 'EvaHsu' THEN  'EvaHsuS' ELSE HOMETEL END HOMETEL  FROM OHEM WHERE HOMETEL=@HOMETEL ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOMETEL", HOMETEL));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["invoicem"];
        }

        public System.Data.DataTable GetOHEMSHIP1()
        {

            SqlConnection MyConnection = globals.shipConnection;
            string sql = "SELECT HOMETEL FROM OHEM WHERE DEPT IN (7) AND ISNULL(TERMDATE,'') ='' ORDER BY HOMETEL";

            SqlDataAdapter da = new SqlDataAdapter(sql, MyConnection);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "RMA_PARAMS");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["RMA_PARAMS"];
        }

 
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Transactions;
using System.Configuration;
using System.Net;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Web.UI;
using System.Collections;
using Microsoft.VisualBasic.Devices;
using System.Net.Mime;
namespace ACME
{
    public partial class RmaSZ : ACME.fmBase1
    {
        string strCn02 = "Data Source=acmesap;Initial Catalog=acmesql05;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string H1 = "";
        string H2 = "";
        private StreamWriter sw;
        Attachment data = null;
        private System.Data.DataTable OrderData;
        string NewFileName = "";
        string F = "";

        public RmaSZ()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            rma_mainSZTableAdapter.Connection = MyConnection;

            rma_InvoiceDSZTableAdapter.Connection = MyConnection;

            rma_PackingListDSZTableAdapter.Connection = MyConnection;

        }

        public override void AfterEdit()
        {
            modifyNameTextBox.Text = fmLogin.LoginID.ToString();
            shippingCodeTextBox.ReadOnly = true;
            receiveDayTextBox.ReadOnly = true;
            boardCountNoTextBox.ReadOnly = true;
            shipToDateTextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
            shipToDateTextBox.ReadOnly = true;
        }
        public override void AfterCancelEdit()
        {
            Control();

        }
        private void Control()
        {
            shippingCodeTextBox.ReadOnly = true;
            receiveDayTextBox.ReadOnly = true;
            boardCountNoTextBox.ReadOnly = true;
            shipToDateTextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
            shipToDateTextBox.ReadOnly = true;



            button8.Enabled = true;
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            checkBox1.Enabled = true;
            checkBox2.Enabled = true;
            checkBox3.Enabled = true;
            checkBox4.Enabled = true;
        }

        public override void AfterAddNew()
        {
            Control();
        }
        public override void EndEdit()
        {
            Control();
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                rm.Rma_InvoiceDSZ.RejectChanges();
                rm.Rma_PackingListDSZ.RejectChanges();
            }
            catch
            {
            }

            return true;
        }
        public void UPDATEJOBNO(string u_jobno, DateTime U_ACME_BackDate1, string u_acme_backqty1, string u_rma_no)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("update octr set u_jobno=@u_jobno,U_ACME_BackDate1=@U_ACME_BackDate1,u_acme_backqty1=@u_acme_backqty1,U_PKind=CASE WHEN U_Rquinity =@u_acme_backqty1  THEN 5 END where u_rma_no=@u_rma_no", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@u_jobno", u_jobno));
            command.Parameters.Add(new SqlParameter("@U_ACME_BackDate1", U_ACME_BackDate1));
            command.Parameters.Add(new SqlParameter("@u_acme_backqty1", u_acme_backqty1));
            command.Parameters.Add(new SqlParameter("@u_rma_no", u_rma_no));
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
        public override void AfterEndEdit()
        {
            try
            {

                if (rma_PackingListDSZDataGridView.Rows.Count > 1)
                {



                    System.Data.DataTable dt4 = GetOrderData4();
                    string d;
                    string f;

                    CalcTotals2();
                    DataRow drw4 = dt4.Rows[0];
                    int g = drw4["PackageNo"].ToString().LastIndexOf("-");
                    if (g == 0)
                    {
                        f = drw4["PackageNo"].ToString();
                    }
                    else
                    {
                        f = drw4["PackageNo"].ToString().Substring(g + 1);
                    }
                    if (add6TextBox.Text == "")
                    {


                        int ss = drw4["cno"].ToString().LastIndexOf("~");
                        if (ss == 0)
                        {
                            d = drw4["cno"].ToString();
                        }
                        else
                        {
                            d = drw4["cno"].ToString().Substring(ss + 1);
                        }

                        if (f != "")
                        {
                            int amountText = Convert.ToInt32(f);
                            string s = f;
                            add6TextBox.Text = new Class1().NumberToString2(amountText, s, d);
                        }
                    }

                    if (!String.IsNullOrEmpty(createDateTextBox.Text))
                    {
                        System.Data.DataTable dt1 = rm.Rma_InvoiceDSZ;
                        try
                        {
                            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                            {
                                DateTime R1 = Convert.ToDateTime(createDateTextBox.Text);

                                DataRow drw = dt1.Rows[i];
                                string aa = drw["RmaNo"].ToString();
                                string bb = drw["shippingcode"].ToString();
                                string InQty = drw["InQty"].ToString();
                                UPDATEJOBNO(bb, R1, InQty, aa);
                            }
                        }
                        catch { }


                    }

                }

                rma_mainSZBindingSource.EndEdit();
                rma_mainSZTableAdapter.Update(rm.Rma_mainSZ);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void CalcTotals2()
        {
            try
            {

                Int32 Quantity = 0;
                decimal NET = 0;
                decimal GROSS = 0;


                int i = this.rma_PackingListDSZDataGridView.Rows.Count - 2;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    if (!String.IsNullOrEmpty(rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Quantity"].Value.ToString()))
                    {
                        int g = rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Quantity"].Value.ToString().LastIndexOf("@");
                        if (g != 0)
                        {
                            Quantity += Convert.ToInt32(rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Quantity"].Value);

                        }
                    }
                    if (!String.IsNullOrEmpty(rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Net"].Value.ToString()))
                    {
                        int U = rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Net"].Value.ToString().LastIndexOf("@");
                        if (U != 0)
                        {

                            NET += Convert.ToDecimal(rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Net"].Value);
                        }
                    }

                    if (!String.IsNullOrEmpty(rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Gross"].Value.ToString()))
                    {

                        int V = rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Gross"].Value.ToString().LastIndexOf("@");
                        if (V != 0)
                        {
                            GROSS += Convert.ToDecimal(rma_PackingListDSZDataGridView.Rows[iRecs].Cells["Gross"].Value);
                        }
                    }


                }

                add4TextBox.Text = Quantity.ToString();
                add5TextBox.Text = NET.ToString();
                add3TextBox.Text = GROSS.ToString();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private System.Data.DataTable GetOrderData4()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select top 1 packageno,seqno,cno from Rma_PackingListDSZ");
            sb.Append(" where shippingcode=@shippingcode  order by cast(seqno as int) desc ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Rma_PackingListD");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        public override void SetInit()
        {

            MyBS = rma_mainSZBindingSource;
            MyTableName = "Rma_mainSZ";
            MyIDFieldName = "ShippingCode";

            //處理複製
            MasterTable = rm.Rma_mainSZ;
            DetailTables = new System.Data.DataTable[] { rm.Rma_InvoiceDSZ };
            DetailBindingSources = new BindingSource[] { rma_InvoiceDSZBindingSource };

        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {
                string NumberName = "";
                if (globals.DBNAME == "達睿生")
                {
                    NumberName = "RMD" + DateTime.Now.ToString("yyyyMMdd");
                }
                else
                {
                    NumberName = "RMS" + DateTime.Now.ToString("yyyyMMdd");
                }

                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;
            kyes = this.shippingCodeTextBox.Text;
            createNameTextBox.Text = fmLogin.LoginID.ToString();

            this.rma_mainSZBindingSource.EndEdit();
            kyes = null;

            receiveDayTextBox.Text = "TRUCK";
            boardCountNoTextBox.Text = "三角";

            buCardcodeTextBox.Text = "進金生";
            dollarsKindCheckBox.Checked = false;
            //TRUCK
        }
        public override void AfterCopy()
        {
            if (kyes == null)
            {
                string NumberName = "RMS" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                this.shippingCodeTextBox.Text = NumberName + AutoNum + "X";
                kyes = this.shippingCodeTextBox.Text;
            }
        }

        public override void FillData()
        {
            try
            {

                rma_mainSZTableAdapter.Fill(rm.Rma_mainSZ, MyID);
                rma_InvoiceDSZTableAdapter.Fill(rm.Rma_InvoiceDSZ, MyID);

                rma_PackingListDSZTableAdapter.Fill(rm.Rma_PackingListDSZ, MyID);

                checkBox5.Checked = false;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public override bool UpdateData()
        {
            bool UpdateData;

            try
            {
                Validate();

                rma_InvoiceDSZBindingSource.MoveFirst();
                for (int i = 0; i <= rma_InvoiceDSZBindingSource.Count - 1; i++)
                {
                    DataRowView row = (DataRowView)rma_InvoiceDSZBindingSource.Current;
                    row["SeqNo"] = i;
                    rma_InvoiceDSZBindingSource.EndEdit();
                    rma_InvoiceDSZBindingSource.MoveNext();
                }


                rma_PackingListDSZBindingSource.MoveFirst();
                for (int i = 0; i <= rma_PackingListDSZBindingSource.Count - 1; i++)
                {
                    DataRowView row2 = (DataRowView)rma_PackingListDSZBindingSource.Current;
                    row2["SeqNo"] = i;
                    rma_PackingListDSZBindingSource.EndEdit();
                    rma_PackingListDSZBindingSource.MoveNext();
                }


                rma_mainSZTableAdapter.Connection.Open();


                rma_mainSZBindingSource.EndEdit();
                rma_InvoiceDSZBindingSource.EndEdit();

                rma_PackingListDSZBindingSource.EndEdit();


                rma_mainSZTableAdapter.Update(rm.Rma_mainSZ);

                rma_InvoiceDSZTableAdapter.Update(rm.Rma_InvoiceDSZ);

                rma_PackingListDSZTableAdapter.Update(rm.Rma_PackingListDSZ);



                UpdateData = true;
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.rma_mainSZTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            pINOTextBox.Text = comboBox1.Text;
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            cFSTextBox.Text = comboBox2.Text;
        }
        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
        }

        private void rma_InvoiceDSZDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (rma_InvoiceDSZDataGridView.Columns[e.ColumnIndex].Name == "InQty" ||
                 rma_InvoiceDSZDataGridView.Columns[e.ColumnIndex].Name == "UnitPrice")
                {
                    decimal iQuantity = 0;
                    decimal iUnitPrice = 0;

                    iQuantity = Convert.ToInt32(this.rma_InvoiceDSZDataGridView.Rows[e.RowIndex].Cells["InQty"].Value);
                    iUnitPrice = Convert.ToDecimal(this.rma_InvoiceDSZDataGridView.Rows[e.RowIndex].Cells["UnitPrice"].Value);
                    this.rma_InvoiceDSZDataGridView.Rows[e.RowIndex].Cells["Amount"].Value = (iQuantity * iUnitPrice).ToString();

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        private void rma_InvoiceDSZDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = rma_InvoiceDSZDataGridView.Rows.Count - 1;
            e.Row.Cells["SeqNo"].Value = iRecs.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuList();

            if (LookupValues != null)
            {

                cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                cardNameTextBox.Text = Convert.ToString(LookupValues[1]);
                add10TextBox.Text = Convert.ToString(LookupValues[6]);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                H1 = "進金生";
            }
            if (radioButton2.Checked)
            {
                H1 = "達睿生";
            }

            RmaNo frm1 = new RmaNo();
            frm1.q1 = H1;
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (H1 == "進金生")
                    {
                        H2 = frm1.q;
                    }
                    if (H1 == "達睿生")
                    {
                        H2 = frm1.q2;
                    }
                    System.Data.DataTable dt1 = GetAR2(H2, H1);

                    System.Data.DataTable dt2 = rm.Rma_InvoiceDSZ;

                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();



                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["RmaNo"] = drw["U_RMA_NO"];
                        drw2["MarkNos"] = drw["U_RMODEL"];
                        drw2["SeqNo"] = "0";
                        drw2["VenderNo"] = drw["U_AUO_RMA_NO"];

                        string K1 = drw["QTY"].ToString();
                        if (ValidateUtils.IsNumeric(K1))
                        {
                            drw2["QTY"] = K1;
                        }
                        else
                        {
                            drw2["QTY"] = "0";
                        }

                        string SIZE = drw["bb"].ToString();

                        drw2["size"] = SIZE;
                        string T1 = drw["aa"].ToString();
                        string T2 = drw["U_RMODEL"].ToString();

                        string SIZE2 = "";
                        int I2 = T2.ToUpper().IndexOf("_OPEN CELL");
                        int I3 = T2.ToUpper().IndexOf("KIT_");
                        int I4 = T2.ToUpper().IndexOf("KIT AD BOARD");
                        int I5 = T2.ToUpper().IndexOf("KIT AD INVERTER");
                        int I6 = T2.ToUpper().IndexOf("KIT DRIVER BOARD");
                        int I7 = T2.ToUpper().IndexOf("KIT INVERTER");
                        int I8 = T2.ToUpper().IndexOf("_T CON");

                        //OpenFrame 
                        int I9 = T2.ToUpper().IndexOf("\"");
                        int I11 = T2.ToUpper().IndexOf("”");
                        int I10 = T2.ToUpper().IndexOf("OPENFRAME");
                        if (I9 != -1)
                        {
                            SIZE2 = T2.Substring(0, I9);
                        }
                        if (I11 != -1)
                        {
                            SIZE2 = T2.Substring(0, I11);
                        }


                        if (I2 != -1)
                        {

                            drw2["MarkNos"] = drw["U_RMODEL"].ToString().Substring(0, I2);
                            drw2["INDescription"] = SIZE + "\" OPEN CELL_AU";


                        }
                        else if (I8 != -1)
                        {
                            drw2["MarkNos"] = drw["U_RMODEL"].ToString().Substring(0, I8);
                            drw2["INDescription"] = "TCON-PCBA_";

                        }
                        else if (I3 != -1)
                        {
                            drw2["MarkNos"] = SIZE2 + "\" KIT";
                            drw2["INDescription"] = SIZE2 + "\" KIT";
                            drw2["DIFF"] = "Y";
                        }
                        else if (I4 != -1)
                        {
                            drw2["MarkNos"] = SIZE2 + "\" KIT_AD Board";
                            drw2["INDescription"] = SIZE2 + "\" KIT_AD Board";
                            drw2["DIFF"] = "Y";
                        }
                        else if (I5 != -1)
                        {
                            drw2["MarkNos"] = SIZE2 + "\" KIT_Inverter";
                            drw2["INDescription"] = SIZE2 + "\" KIT_Inverter";
                            drw2["DIFF"] = "Y";
                        }
                        else if (I6 != -1)
                        {
                            drw2["MarkNos"] = SIZE2 + "\" KIT_Driver Board";
                            drw2["INDescription"] = SIZE2 + "\" KIT_Driver Board";
                            drw2["DIFF"] = "Y";
                        }
                        else if (I7 != -1)
                        {
                            drw2["MarkNos"] = SIZE2 + "\" KIT_Inverter";
                            drw2["INDescription"] = SIZE2 + "\" KIT_Inverter";
                            drw2["DIFF"] = "Y";
                        }
                        else if (I10 != -1)
                        {

                            drw2["MarkNos"] = SIZE2 + "\" OpenFrame";
                            drw2["INDescription"] = "OFD_" + SIZE2 + "\" OpenFrame LCD Monitor";
                            drw2["DIFF"] = "Y";
                        }
                        else
                        {
                            drw2["MarkNos"] = drw["U_RMODEL"];
                            drw2["INDescription"] = T1;
                        }
                        string U_Rquinity = drw["U_Rquinity"].ToString();

                        if (!String.IsNullOrEmpty(U_Rquinity))
                        {
                            int g = U_Rquinity.IndexOf("+");
                            int t = U_Rquinity.LastIndexOf("+");
                            string h;
                            string s;


                            if (g == -1)
                            {
                                s = U_Rquinity;
                                drw2["InQty"] = s;
                            }
                            else
                            {
                                s = U_Rquinity.Substring(g + 1);

                                try
                                {

                                    if (U_Rquinity.Substring(3, 1) != "+")
                                    {
                                        h = U_Rquinity.Substring(0, 2);
                                    }
                                    else
                                    {
                                        h = U_Rquinity.Substring(0, 1);
                                    }
                                    int a = Convert.ToInt16(s.ToString());
                                    int b = Convert.ToInt16(h.ToString());
                                    drw2["InQty"] = (a + b).ToString();
                                }
                                catch (Exception ex)
                                {
                                    h = U_Rquinity.Substring(0, 1);
                                    drw2["InQty"] = h.ToString();
                                }

                            }
                        }

                        drw2["Grade"] = drw["U_Rgrade"];
                        drw2["InvoiceNo_seq"] = drw["U_Rver"];
                        drw2["CodeName"] = drw["U_cusname_s"];
                        string VenderNo = drw["U_AUO_RMA_NO"].ToString();
                        drw2["VENDER"] = drw["U_Rvender"];
                        dt2.Rows.Add(drw2);
                    }
                    for (int j = 0; j <= rma_InvoiceDSZDataGridView.Rows.Count - 2; j++)
                    {
                        rma_InvoiceDSZDataGridView.Rows[j].Cells[0].Value = j.ToString();
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {

            }
        }

        public System.Data.DataTable GetAR2(string DocEntry, string q1)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string aa = '"'.ToString();
            if (q1 == "達睿生")
            {
                MyConnection = new SqlConnection(strCn02);
            }
            string sql = "select U_RMA_NO,U_AUO_RMA_NO,U_Rgrade,U_Rmodel,U_cusname_s,U_Rver,U_Rmodel,U_Rvender,U_Rquinity,aa=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1)+'" + aa + "'+'TFT LCD MODULE' END,bb=case substring(U_Rmodel,4,1) when 0 then substring(U_Rmodel,2,2) ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1) END,U_RQUINITY QTY from octr where Contractid IN (" + DocEntry + ") ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
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

        private void button10_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt1 = Getinvoiced(shippingCodeTextBox.Text);
                System.Data.DataTable dt2 = rm.Rma_PackingListDSZ;

                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();

                    string MODEL = drw["MarkNos"].ToString();
                    string VER = drw["InvoiceNo_seq"].ToString();
                    if (drw["INDescription"].ToString().ToUpper().IndexOf("OPEN") != -1 && drw["INDescription"].ToString().ToUpper().IndexOf("CELL") != -1)
                    {

                        MODEL = "O" + MODEL;
                    }

                    System.Data.DataTable J1 = GetNET(MODEL, VER);
                    string NET = "";
                    if (J1.Rows.Count > 0)
                    {
                        decimal H1 = 0;
                        string H2 = J1.Rows[0][0].ToString();
                        decimal QTY2 = 0;
                        string QTY = drw["InQty"].ToString().Replace("@", "");
                        decimal number3 = 0;
                        int number4 = 0;
                        bool canConvert = int.TryParse(QTY, out number4);
                        bool canConvert2 = decimal.TryParse(H2, out number3);
                        if (drw["INDescription"].ToString().ToUpper().IndexOf("KIT") == -1 && drw["INDescription"].ToString().ToUpper().IndexOf("TCON") == -1)
                        {
                            if (canConvert == true)
                            {
                                QTY2 = Convert.ToDecimal(QTY);

                                if (canConvert2 == true)
                                {
                                    H1 = Convert.ToDecimal(H2);
                                    NET = (H1 * QTY2).ToString().Replace(".0", "");
                                }
                            }
                        }
                    }

                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = drw["SeqNo"];
                    drw2["DescGoods"] = (i + 1).ToString() + ")" + drw["INDescription"];
                    drw2["model"] = drw["MarkNos"];
                    drw2["Ver"] = drw["InvoiceNo_seq"];
                    drw2["Quantity"] = drw["InQty"];
                    drw2["RmaNo"] = drw["RmaNo"];
                    drw2["VenderNo"] = drw["VenderNo"];
                    drw2["CardName"] = drw["CodeName"];
                    drw2["Net"] = NET;
                    drw2["DIFF"] = drw["DIFF"];
                    drw2["QTY"] = drw["QTY"];
                    string N1 = drw["InQty"].ToString();
                    string N2 = drw["QTY"].ToString();
                    if (String.IsNullOrEmpty(N1))
                    {
                        N1 = "0";
                    }
                    if (String.IsNullOrEmpty(N2))
                    {
                        N2 = "0";
                    }
                    int a = Convert.ToInt16(N1);
                    int b = Convert.ToInt16(N2);
                    drw2["NQTY"] = (b - a).ToString();
                    //CardName
                    dt2.Rows.Add(drw2);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public static System.Data.DataTable Getinvoiced(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select * from rma_invoicedSZ where shippingcode=@shippingcode ";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
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




        private System.Data.DataTable GetOrderData31(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                                                    SELECT c.ADD8 DOLAR,'Ref No:'+c.ADD9 REF,b.SHIPPINGCODE SHIPPINGCODE,b.RMANO RMANO,c.add7+' PLTS' as PLT,Convert(varchar(10),Getdate(),111) as DATE,c.[receiveDay] SHIPBY,c.[add10] as SHIPTO,c.[receivePlace] SHIPFROM,");
            sb.Append("                                                         c.[add5] as TNET,c.[add3] TGRO,c.[goalPlace] SHIPTO2,c.[add9] BILLTO");
            sb.Append("                                                           ,c.[add4] TQTY,c.[add6] TOTAL,b.[PackageNo] PALNO,b.[CNo],");
            sb.Append(" CASE WHEN B.DIFF='Y' THEN b.[DESCGOODS] ");
            sb.Append("  ELSE b.[DESCGOODS]+ char(10) + char(13)+b.[model]+'  V.'+isnull(b.ver,'') END DES");
            sb.Append("                                                           ,b.[Quantity]  QTY ,b.[Net] as NET ,cast(b.[Gross] as varchar) as GRO ,b.[MeasurmentCM] CM,C.CREATENAME USERS  ");
            sb.Append("                                                      from   [RMA_PackingListDSZ] as b ");
            sb.Append("                                                        left join RMA_mainSZ as c on (b.shippingcode=c.shippingcode)");
            sb.Append("   where b.shippingcode=@shippingcode  ");
            sb.Append("  ORDER BY CAST(seqno AS INT)  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData31MULTI(string SHIPPINGCODE, string COMPANY)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            if (COMPANY == "進金生")
            {

                sb.Append("                                                    SELECT c.ADD8 DOLAR,'Ref No:'+c.ADD9 REF,b.SHIPPINGCODE SHIPPINGCODE,b.RMANO RMANO,c.add7+' PLTS' as PLT,Convert(varchar(10),Getdate(),111) as DATE,c.[receiveDay] SHIPBY,c.[add10] as SHIPTO,c.[receivePlace] SHIPFROM,");
                sb.Append("                                                         c.[add5] as TNET,c.[add3] TGRO,c.[goalPlace] SHIPTO2,c.[add9] BILLTO");
                sb.Append("                                                           ,c.[add4] TQTY,c.[add6] TOTAL,b.[PackageNo] PALNO,b.[CNo],");
                sb.Append(" CASE WHEN B.DIFF='Y' THEN b.[DESCGOODS] ");
                sb.Append("  ELSE b.[DESCGOODS]+ char(10) + char(13)+b.[model]+'  V.'+isnull(b.ver,'') END DES");
                sb.Append("                                                           ,b.[Quantity]  QTY ,b.[Net] as NET ,cast(b.[Gross] as varchar) as GRO ,b.[MeasurmentCM] CM,C.CREATENAME USERS  ");
                sb.Append("                                                      from   ACMESQLSP.DBO.[RMA_PackingListDSZ] as b ");
                sb.Append("                                                        left join ACMESQLSP.DBO.RMA_mainSZ as c on (b.shippingcode=c.shippingcode)");
                sb.Append("   where b.shippingcode=@shippingcode  ");
                sb.Append("  ORDER BY CAST(seqno AS INT)  ");
            }
            if (COMPANY == "達睿生")
            {

                sb.Append("                                                    SELECT c.ADD8 DOLAR,'Ref No:'+c.ADD9 REF,b.SHIPPINGCODE SHIPPINGCODE,b.RMANO RMANO,c.add7+' PLTS' as PLT,Convert(varchar(10),Getdate(),111) as DATE,c.[receiveDay] SHIPBY,c.[add10] as SHIPTO,c.[receivePlace] SHIPFROM,");
                sb.Append("                                                         c.[add5] as TNET,c.[add3] TGRO,c.[goalPlace] SHIPTO2,c.[add9] BILLTO");
                sb.Append("                                                           ,c.[add4] TQTY,c.[add6] TOTAL,b.[PackageNo] PALNO,b.[CNo],");
                sb.Append(" CASE WHEN B.DIFF='Y' THEN b.[DESCGOODS] ");
                sb.Append("  ELSE b.[DESCGOODS]+ char(10) + char(13)+b.[model]+'  V.'+isnull(b.ver,'') END DES");
                sb.Append("                                                           ,b.[Quantity]  QTY ,b.[Net] as NET ,cast(b.[Gross] as varchar) as GRO ,b.[MeasurmentCM] CM,C.CREATENAME USERS  ");
                sb.Append("                                                      from   ACMESQLSPDRS.DBO.[RMA_PackingListDSZ] as b ");
                sb.Append("                                                        left join ACMESQLSPDRS.DBO.RMA_mainSZ as c on (b.shippingcode=c.shippingcode)");
                sb.Append("   where b.shippingcode=@shippingcode  ");
                sb.Append("  ORDER BY CAST(seqno AS INT)  ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT boatCompany SHIPTO,''''+PackageNo PLNO,''''+CNo CNO,T1.RMANO,MODEL,VER,T1.QUANTITY QTY,T1.VENDERNO,T1.RMANO,T0.add6 TOTAL,T1.SEQNO,T1.DescGoods [DESC] ");
            sb.Append("  FROM dbo.Rma_mainSZ T0");
            sb.Append(" LEFT JOIN dbo.Rma_PackingListDSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.SHIPPINGCODE = @SHIPPINGCODE ORDER BY CAST(T1.SEQNO AS INT)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3CREATE(string CREATEDATE, string COMPANY)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            if (COMPANY == "進金生")
            {
                sb.Append("               SELECT boatCompany SHIPTO,''''+PackageNo PLNO,''''+CNo CNO,T1.RMANO,MODEL,VER,T1.QUANTITY QTY,T1.VENDERNO,T1.RMANO,T0.add6 TOTAL,T1.SEQNO,T1.DescGoods [DESC]  ");
                sb.Append("                ,T0.SHIPPINGCODE JOBNO FROM ACMESQLSP.DBO.Rma_mainSZ T0 ");
                sb.Append("               LEFT JOIN ACMESQLSP.DBO.Rma_PackingListDSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
                sb.Append(" WHERE T0.CREATEDATE = @CREATEDATE ORDER BY T0.SHIPPINGCODE,CAST(T1.SEQNO AS INT)");
            }
            if (COMPANY == "達睿生")
            {
                sb.Append("               SELECT boatCompany SHIPTO,''''+PackageNo PLNO,''''+CNo CNO,T1.RMANO,MODEL,VER,T1.QUANTITY QTY,T1.VENDERNO,T1.RMANO,T0.add6 TOTAL,T1.SEQNO,T1.DescGoods [DESC]  ");
                sb.Append("                ,T0.SHIPPINGCODE JOBNO FROM ACMESQLSPDRS.DBO.Rma_mainSZ T0 ");
                sb.Append("               LEFT JOIN ACMESQLSPDRS.DBO.Rma_PackingListDSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
                sb.Append(" WHERE T0.CREATEDATE = @CREATEDATE ORDER BY T0.SHIPPINGCODE,CAST(T1.SEQNO AS INT)");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CREATEDATE", CREATEDATE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3G()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT T1.SHIPPINGCODE,RMANO,CODENAME CARDNAME,InvoiceNo_seq VER,QTY,INQTY FROM rma_INVOICEDSZ T0 LEFT JOIN RMA_MAINSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE  REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB ORDER BY T1.SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text.Trim()));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text.Trim()));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3G2(string RMANO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT TOP 1  INQTY FROM rma_INVOICEDSZ WHERE RMANO=@RMANO ORDER BY SHIPPINGCODE DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@RMANO", RMANO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetACMERET(string U_RMA_NO)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT TOP 1 U_ACME_BACKQTY1 FROM OCTR WHERE U_RMA_NO=@U_RMA_NO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3GS(string FD)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            if (FD == "進金生")
            {
                sb.Append("SELECT T1.SHIPPINGCODE,''''+PackageNo PackageNo,T0.CARDNAME,CNo,Gross,MeasurmentCM FROM  ACMESQLSP.DBO.rma_PackingListDSZ T0 LEFT JOIN  ACMESQLSP.DBO.RMA_MAINSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE  REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB AND ISNULL(CNo,'') <> '' ORDER BY T1.SHIPPINGCODE");
            }
            if (FD == "達睿生")
            {
                sb.Append("SELECT T1.SHIPPINGCODE,''''+PackageNo PackageNo,T0.CARDNAME,CNo,Gross,MeasurmentCM FROM  ACMESQLSPDRS.DBO.rma_PackingListDSZ T0 LEFT JOIN  ACMESQLSPDRS.DBO.RMA_MAINSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE  REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB AND ISNULL(CNo,'') <> '' ORDER BY T1.SHIPPINGCODE");
            }
            if (FD == "進金生達睿生")
            {
                sb.Append("SELECT T1.SHIPPINGCODE,''''+PackageNo PackageNo,T0.CARDNAME,CNo,Gross,MeasurmentCM FROM  ACMESQLSP.DBO.rma_PackingListDSZ T0 LEFT JOIN  ACMESQLSP.DBO.RMA_MAINSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE  REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB AND ISNULL(CNo,'') <> '' ");
                sb.Append("UNION ALL ");
                sb.Append("SELECT T1.SHIPPINGCODE,''''+PackageNo PackageNo,T0.CARDNAME,CNo,Gross,MeasurmentCM FROM  ACMESQLSPDRS.DBO.rma_PackingListDSZ T0 LEFT JOIN  ACMESQLSPDRS.DBO.RMA_MAINSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE  REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB AND ISNULL(CNo,'') <> '' ORDER BY T1.SHIPPINGCODE");
            }


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text.Trim()));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text.Trim()));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3F(string FD)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            if (FD == "進金生")
            {
                sb.Append(" SELECT SHIPPINGCODE,boatCompany CARDNAME,'進金生' COMPANY  FROM ACMESQLSP.DBO.Rma_mainSZ  WHERE REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB ");
            }
            if (FD == "達睿生")
            {
                sb.Append(" SELECT SHIPPINGCODE,boatCompany CARDNAME,'達睿生' COMPANY  FROM ACMESQLSPDRS.DBO.Rma_mainSZ  WHERE REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB ");
            }
            if (FD == "進金生達睿生")
            {
                sb.Append(" SELECT SHIPPINGCODE,boatCompany CARDNAME,'進金生' COMPANY  FROM ACMESQLSP.DBO.Rma_mainSZ  WHERE REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB ");
                sb.Append(" UNION ALL ");
                sb.Append(" SELECT SHIPPINGCODE,boatCompany CARDNAME,'達睿生' COMPANY  FROM ACMESQLSPDRS.DBO.Rma_mainSZ  WHERE REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3FCREATE()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT CREATEDATE CREATEDATE,'進金生' COMPANY  FROM ACMESQLSP.DBO.Rma_mainSZ   WHERE REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB ");
            sb.Append(" UNION ALL ");
            sb.Append(" SELECT DISTINCT CREATEDATE CREATEDATE,'達睿生' COMPANY  FROM ACMESQLSPDRS.DBO.Rma_mainSZ   WHERE REPLACE(REPLACE(REPLACE(createDate,'/',''),'.',''),'-','') BETWEEN @AA AND @BB ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMODEL(string U_RMA_NO)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT U_RMODEL  FROM OCTR T1 WHERE T1.U_RMA_NO=@U_RMA_NO ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData31S(string SEQNO)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT QUANTITY  FROM dbo.Rma_PackingListDSZ WHERE SHIPPINGCODE = @SHIPPINGCODE AND SEQNO=@SEQNO");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@SEQNO", SEQNO));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3CNO(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT COUNT(*) FROM  dbo.Rma_PackingListDSZ  WHERE SHIPPINGCODE = @SHIPPINGCODE     ");


            sb.Append(" AND ISNULL(CNO,'') <> ''   ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3CNOCREATE(string CREATEDATE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT COUNT(*) FROM  dbo.Rma_PackingListDSZ T0 LEFT JOIN Rma_mainSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  WHERE CREATEDATE = @CREATEDATE     ");


            sb.Append(" AND ISNULL(CNO,'') <> ''   ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CREATEDATE", CREATEDATE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3CNO1(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ''''+CNO FROM  dbo.Rma_PackingListDSZ  WHERE SHIPPINGCODE = @SHIPPINGCODE  ");

            sb.Append(" AND ISNULL(CNO,'') <> ''   ");


            sb.Append("  ORDER BY CAST(SEQNO AS INT)  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3CNO1CREATE(string CREATEDATE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ''''+CNO,T0.SHIPPINGCODE,T0.CARDNAME FROM dbo.Rma_PackingListDSZ T0 LEFT JOIN Rma_mainSZ T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)  WHERE CREATEDATE = @CREATEDATE  ");

            sb.Append(" AND ISNULL(CNO,'') <> ''   ");


            sb.Append("  ORDER BY T0.SHIPPINGCODE,CAST(SEQNO AS INT)  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CREATEDATE", CREATEDATE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3CNO2(string SHIPPINGCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("SELECT ''''+CNO FROM  dbo.Rma_PackingListDSZ  WHERE SHIPPINGCODE = @SHIPPINGCODE  ");

            sb.Append(" AND ISNULL(CNO,'') <> ''   ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void GetExcelProduct4(string ExcelFile, string PRINT, System.Data.DataTable dt,  string P1, string FLAG)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(ExcelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;
            Microsoft.Office.Interop.Excel.Worksheet SheetTemplate = null;
            //第一個當作範本
            SheetTemplate = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.ActiveSheet;


           

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
            //excelSheet.Name = shippingCodeTextBox.Text;

            //取得 Excel 的使用區域
            int iRowCnt = SheetTemplate.UsedRange.Cells.Rows.Count;
            int iColCnt = SheetTemplate.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;

            string SHIPPINGCODE = "";
            string COMPANY = "";
            string CARDNAME = "";
            Microsoft.Office.Interop.Excel.Range range = null;
            try 
            {
                foreach (DataRow row in dt.Rows)
                {
                    SHIPPINGCODE = row["SHIPPINGCODE"].ToString();
                    CARDNAME = row["CARDNAME"].ToString();
                    COMPANY = row["COMPANY"].ToString();

                    OrderData = GetOrderData31MULTI(SHIPPINGCODE, COMPANY);


                    SheetTemplate.Copy(Type.Missing, excelBook.Sheets[excelBook.Sheets.Count]);

                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets[excelBook.Sheets.Count];
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets[excelBook.Sheets.Count];
                    excelSheet.Name = CARDNAME + "-" + SHIPPINGCODE;


                    string sTemp = string.Empty;
                    string FieldValue = string.Empty;
                    string FieldValue1 = string.Empty;
                    bool IsDetail = false;
                    int DetailRow = 0;
                    int DetailRow1 = 0;

                    if (FLAG == "Y")
                    {
                        //                excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
                        //Microsoft.Office.Core.MsoTriState.msoTrue, Convert.ToInt16(textBoxF1.Text), Convert.ToInt16(textBoxF2.Text), Convert.ToInt16(textBoxF3.Text), Convert.ToInt16(textBoxF4.Text));
                        excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
        Microsoft.Office.Core.MsoTriState.msoTrue, 395, 625, 150, 60);
                    }

                    for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                    {

                        for (int iField = 1; iField <= iColCnt; iField++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            if (CheckSerial(sTemp, ref FieldValue))
                            {
                                range.Value2 = FieldValue;
                            }
                            if (IsDetailRow(sTemp))
                            {
                                IsDetail = true;
                                DetailRow = iRecord;
                                DetailRow1 = 9;
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
                                // range.Select();
                                sTemp = (string)range.Text;
                                sTemp = sTemp.Trim();

                                FieldValue = "";
                                SetRow(aRow, sTemp, ref FieldValue);

                                range.Value2 = FieldValue;


                            }
                            DetailRow++;
                        }
                    }
                }
            }
            catch (Exception ex)
            { 

            }
            finally
            {

                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                 DateTime.Now.ToString("yyyy.MM.dd") + " PACK.xls";

                SetMsg(CARDNAME + SHIPPINGCODE + "PACK.xls");
                try
                {
                    SheetTemplate.Delete();

                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

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

                string Msg = string.Empty;


                if (PRINT != "N")
                {
                    System.Diagnostics.Process.Start(NewFileName);
                }



            }
        }


        private void GetExcelProductCREATE(string ExcelFile, string PRINT, string CREATEDATE)
        {

            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;
            object SelectCell = null;
            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
            excelSheet.Name = shippingCodeTextBox.Text;

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 9;
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
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }
                object Cell_From;
                object Cell_To;
                object FixCell;
                object FixCell2;

                string ID1 = "";
                string ID2 = "";
                string ID3 = "";
                string ID4 = "";
                int iRowCnt3 = excelSheet.UsedRange.Cells.Rows.Count;
                string numString = GetOrderData3CNOCREATE(CREATEDATE).Rows[0][0].ToString();

                int number1 = 0;
                bool canConvert = int.TryParse(numString, out number1);
                if (canConvert == true)
                {
                    Cell_From = "A1";
                    Cell_To = "K" + Convert.ToString(iRowCnt3 + 1);
                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);
                    range.Select();
                    System.Data.DataTable K2 = GetOrderData3CNO1CREATE(CREATEDATE);
                    for (int aRow = 3; aRow <= iRowCnt3; aRow++)
                    {

                        string N1 = K2.Rows[0][0].ToString();
                        string N2 = K2.Rows[0][1].ToString();
                        string N3 = K2.Rows[0][2].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 12]);
                        range.Select();
                        range.Value2 = N1.ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 13]);
                        range.Select();
                        range.Value2 = N2.ToString();

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 4]);
                        range.Select();
                        range.Value2 = N3.ToString();
                    }
                    int F1 = Convert.ToInt16(numString);
                    for (int i = 1; i <= F1 - 1; i++)
                    {

                        int COPY = ((iRowCnt3 + 1) * i) + 1;

                        int COPY2 = ((iRowCnt3 + 1) * (i + 1)) + 1;
                        FixCell = "A" + Convert.ToString(COPY);
                        FixCell2 = "K" + Convert.ToString(COPY2);
                        range = excelSheet.get_Range(FixCell, FixCell2);
                        range.Select();
                        excelSheet.Paste(oMissing, oMissing);
                        System.Data.DataTable K1 = GetOrderData3CNO1CREATE(CREATEDATE);
                        string N1 = K1.Rows[i][0].ToString();
                        string N2 = K1.Rows[i][1].ToString();
                        string N3 = K1.Rows[i][2].ToString();
                        for (int aRow = COPY; aRow <= COPY2; aRow++)
                        {


                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 12]);
                            range.Select();
                            range.Value2 = N1.ToString();

                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 13]);
                            range.Select();
                            range.Value2 = N2.ToString();

                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[COPY, 4]);
                        range.Select();
                        range.Value2 = N3.ToString();
                    }

                    int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;

                    string FLAG = "";
                    for (int aRow = 1; aRow <= iRowCnt2; aRow++)
                    {
                        if (FLAG == "Y")
                        {

                            aRow = aRow - 1;
                            FLAG = "";
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID1 = sTemp;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 12]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID2 = sTemp;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 11]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID3 = sTemp;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 13]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID4 = sTemp;


                        if (!String.IsNullOrEmpty(ID1))
                        {
                            if (ID1 != ID2)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 1]);
                                range.Select();
                                range.EntireRow.Delete(XlDirection.xlDown);
                                FLAG = "Y";

                            }
                            else if (ID3 != ID4)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 1]);
                                range.Select();
                                range.EntireRow.Delete(XlDirection.xlDown);
                                FLAG = "Y";

                            }
                        }

                    }


                }
                else
                {
                    return;
                }


            }
            finally
            {



                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" + "还货唛头_" +
CREATEDATE + ".xls";
                SetMsg("还货唛头_" +
CREATEDATE + ".xls");
                try
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 13]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 12]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 11]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 10]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

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

                string Msg = string.Empty;

                if (PRINT == "Y")
                {
                    System.Diagnostics.Process.Start(NewFileName);
                }



            }
        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private void GetExcelProduct3(string ExcelFile, string PRINT, string SHIPPINGCODE, string CARDNAME, string P1, string FLAG)
        {
            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
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
            excelSheet.Name = shippingCodeTextBox.Text;

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;

                if (FLAG == "Y")
                {

                    excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
    Microsoft.Office.Core.MsoTriState.msoTrue, 395, 625, 150, 60);
                }

                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 9;
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
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }




            }
            finally
            {

                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  CARDNAME + SHIPPINGCODE + "PACK.xls";

                SetMsg(CARDNAME + SHIPPINGCODE + "PACK.xls");
                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

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

                string Msg = string.Empty;


                if (PRINT != "N")
                {
                    System.Diagnostics.Process.Start(NewFileName);
                }



            }
        }
        private void GetExcelProduct(string ExcelFile, string PRINT, string SHIPPINGCODE, string J1, string CARDNAME)
        {

            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;
            object SelectCell = null;
            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
            excelSheet.Name = shippingCodeTextBox.Text;

            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            // progressBar1.Maximum = iRowCnt;


            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();

                        if (CheckSerial(sTemp, ref FieldValue))
                        {
                            range.Value2 = FieldValue;
                        }

                        //檢查是不是 Detail Row
                        //要先作完所有 Master 之後再去作 Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            DetailRow1 = 9;
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
                            // range.Select();
                            sTemp = (string)range.Text;
                            sTemp = sTemp.Trim();

                            FieldValue = "";
                            SetRow(aRow, sTemp, ref FieldValue);

                            range.Value2 = FieldValue;


                        }

                        DetailRow++;
                    }

                }
                object Cell_From;
                object Cell_To;
                object FixCell;
                object FixCell2;

                string ID1 = "";
                string ID2 = "";

                int iRowCnt3 = excelSheet.UsedRange.Cells.Rows.Count;
                string numString = GetOrderData3CNOCREATE(SHIPPINGCODE).Rows[0][0].ToString();

                int number1 = 0;
                bool canConvert = int.TryParse(numString, out number1);
                if (canConvert == true)
                {
                    Cell_From = "A1";
                    Cell_To = "J" + Convert.ToString(iRowCnt3 + 6);
                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);
                    range.Select();
                    for (int aRow = 3; aRow <= iRowCnt; aRow++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 11]);
                        range.Select();
                        string N1 = GetOrderData3CNO1(SHIPPINGCODE).Rows[0][0].ToString();
                        range.Value2 = N1.ToString();
                    }
                    int F1 = Convert.ToInt16(numString);
                    for (int i = 1; i <= F1 - 1; i++)
                    {

                        int COPY = ((iRowCnt3 + 6) * i) + 1;

                        int COPY2 = ((iRowCnt3 + 6) * (i + 1)) + 1;
                        FixCell = "A" + Convert.ToString(COPY);
                        FixCell2 = "J" + Convert.ToString(COPY2);
                        range = excelSheet.get_Range(FixCell, FixCell2);
                        range.Select();
                        excelSheet.Paste(oMissing, oMissing);

                        for (int aRow = COPY; aRow <= COPY2; aRow++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 11]);
                            range.Select();
                            string N2 = GetOrderData3CNO1(SHIPPINGCODE).Rows[i][0].ToString();
                            range.Value2 = N2.ToString();

                        }


                    }

                    int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;

                    string FLAG = "";
                    for (int aRow = 1; aRow <= iRowCnt2; aRow++)
                    {
                        if (FLAG == "Y")
                        {

                            aRow = aRow - 1;
                            FLAG = "";
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID1 = sTemp;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 11]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID2 = sTemp;

                        if (!String.IsNullOrEmpty(ID1))
                        {
                            if (ID1 != ID2)
                            {
                                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 1]);
                                range.Select();
                                range.EntireRow.Delete(XlDirection.xlDown);
                                FLAG = "Y";

                            }
                        }

                    }


                }
                else
                {
                    return;
                }


            }
            finally
            {



                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" + "还货唛头_" +
CARDNAME + " " + J1 + ".xls";
                SetMsg("还货唛头_" +
CARDNAME + " " + J1 + ".xls");
                try
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 11]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 10]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

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

                string Msg = string.Empty;

                if (PRINT == "Y")
                {
                    System.Diagnostics.Process.Start(NewFileName);
                }



            }
        }

        private void SetRow(int iRow, string sData, ref string FieldValue)
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
        private bool IsDetailRow(string sData)
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

        private bool CheckSerial(string sData, ref string FieldValue)
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

        private void rma_PackingListDSZDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            //Quantity
            int iRecs;
            iRecs = rma_PackingListDSZDataGridView.Rows.Count - 1;
            e.Row.Cells["dataGridViewTextBoxColumn18"].Value = iRecs.ToString();
            e.Row.Cells["NQTY"].Value = "0";
            e.Row.Cells["QTY1"].Value = "0";
            e.Row.Cells["Quantity"].Value = "0";
        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("SHIPTO", typeof(string));
            dt.Columns.Add("PLNO", typeof(string));
            dt.Columns.Add("CNO", typeof(string));
            dt.Columns.Add("RMANO", typeof(string));
            dt.Columns.Add("MODEL", typeof(string));
            dt.Columns.Add("VER", typeof(string));
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("CNO2", typeof(string));
            dt.Columns.Add("VENDERNO", typeof(string));
            dt.Columns.Add("TOTAL", typeof(string));
            dt.Columns.Add("JOBNO", typeof(string));
            dt.Columns.Add("DESC", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombineG()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("RMA", typeof(string));
            dt.Columns.Add("客户简称", typeof(string));
            dt.Columns.Add("Model", typeof(string));
            dt.Columns.Add("Ver", typeof(string));
            dt.Columns.Add("原退数量", typeof(string));
            dt.Columns.Add("还货数量(Qty)", typeof(string));
            dt.Columns.Add("未还数量", typeof(string));
            dt.Columns.Add("Remark", typeof(string));
            dt.Columns.Add("JOB NO", typeof(string));

            return dt;
        }
        System.Data.DataTable DT(string SHIPPINGCODE, string CREATEDATE, string COMPANY)
        {
            System.Data.DataTable dtCost = MakeTableCombine();
            System.Data.DataTable dt = null;
            if (CREATEDATE == "")
            {
                dt = GetOrderData3(SHIPPINGCODE);
            }
            else
            {
                dt = GetOrderData3CREATE(CREATEDATE, COMPANY);
            }
            string id1x = "";
            string id = "";
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                DataRow dd = dt.Rows[i];
                dr = dtCost.NewRow();

                dr["SHIPTO"] = dd["SHIPTO"].ToString();
                dr["PLNO"] = dd["PLNO"].ToString();

                id = dd["CNO"].ToString();

                if (id != "")
                {
                    id1x = id;
                }
                if (id == "")
                {
                    id = id1x;
                }

                dr["CNO"] = id;
                string RMANO = dd["RMANO"].ToString().Trim();
                dr["RMANO"] = RMANO;

                System.Data.DataTable J1 = GetMODEL(RMANO);
                if (J1.Rows.Count > 0)
                {
                    dr["MODEL"] = J1.Rows[0][0].ToString().Trim();
                }
                else
                {
                    dr["MODEL"] = dd["MODEL"].ToString().Trim();
                }
                dr["VER"] = dd["VER"].ToString();
                int Q1 = dd["QTY"].ToString().IndexOf("@");
                if (Q1 != -1)
                {
                    int g1 = Convert.ToInt16(dd["SEQNO"].ToString()) + 1;

                    System.Data.DataTable T1 = GetOrderData31S(g1.ToString());
                    if (T1.Rows.Count > 0)
                    {
                        dr["QTY"] = T1.Rows[0][0].ToString();
                    }
                    else
                    {
                        dr["QTY"] = dd["QTY"].ToString();
                    }
                }
                else
                {
                    dr["QTY"] = dd["QTY"].ToString();
                }

                dr["CNO2"] = id;
                dr["VENDERNO"] = dd["VENDERNO"].ToString();
                dr["TOTAL"] = dd["TOTAL"].ToString();
                if (CREATEDATE == "")
                {
                    dr["JOBNO"] = "JOB NO.:" + shippingCodeTextBox.Text;
                }
                else
                {
                    dr["JOBNO"] = dd["JOBNO"].ToString();
                }

                dr["DESC"] = dd["DESC"].ToString();
                if (!String.IsNullOrEmpty(dd["MODEL"].ToString()))
                {
                    dtCost.Rows.Add(dr);
                }

            }


            return dtCost;
        }

        System.Data.DataTable DTG()
        {

            System.Data.DataTable dtCost = MakeTableCombineG();
            System.Data.DataTable dt = GetOrderData3G();

            DataRow dr = null;
            int n;
            string JOBNO2 = "";
            int Q1 = 0;
            int Q2 = 0;
            int Q3 = 0;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                DataRow dd = dt.Rows[i];
                dr = dtCost.NewRow();
                string T1 = dd["RMANO"].ToString().Trim();
                string JOBNO = dd["SHIPPINGCODE"].ToString().Trim();
                dr["RMA"] = T1;
                string CARDNAME = dd["CARDNAME"].ToString();
                int G = 0;
                int G2 = 0;

                System.Data.DataTable dt1 = GetACMERET(T1);
                System.Data.DataTable dt12 = GetOrderData3G2(T1);
                if (dt1.Rows.Count > 0)
                {
                    string F = dt1.Rows[0][0].ToString();
                    if (int.TryParse(F, out n))
                    {
                        G = Convert.ToInt16(F);
                    }
                }
                if (dt12.Rows.Count > 0)
                {
                    string F = dt12.Rows[0][0].ToString();
                    if (int.TryParse(F, out n))
                    {
                        G2 = Convert.ToInt16(F);
                    }
                }
                if (JOBNO2 != JOBNO)
                {
                    dr["JOB NO"] = JOBNO;
                    dr["客户简称"] = CARDNAME;
                }
                dr["Ver"] = dd["VER"].ToString();
                string INQTY = dd["INQTY"].ToString().Trim();
                string QTY = dd["QTY"].ToString().Trim();
                dr["原退数量"] = QTY;
                dr["还货数量(Qty)"] = G2.ToString();

                if (int.TryParse(INQTY, out n) && int.TryParse(QTY, out n))
                {
                    int G3 = Convert.ToInt16(QTY);
                    dr["未还数量"] = G3 - G - G2;


                    if (G3 == G)
                    {
                        dr["未还数量"] = 0;
                    }
                    else
                    {
                        Q3 += G3 - G - G2;
                    }
                }
                if (int.TryParse(QTY, out n))
                {
                    Q1 += Convert.ToInt16(QTY);
                }
                Q2 += G2;
                System.Data.DataTable J1 = GetMODEL(T1);
                if (J1.Rows.Count > 0)
                {
                    dr["Model"] = J1.Rows[0][0].ToString().Trim();
                }

                JOBNO2 = JOBNO;
                dtCost.Rows.Add(dr);


            }

            dr = dtCost.NewRow();
            dr["Ver"] = "TTL:";
            dr["原退数量"] = Q1;
            dr["还货数量(Qty)"] = Q2;
            dr["未还数量"] = Q3;
            dtCost.Rows.Add(dr);

            return dtCost;
        }


        private void RmaSZ_Load(object sender, EventArgs e)
        {
            Control();
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            checkedListBox1.SetItemChecked(0, true);
        }



        private void dollarsKindCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (dollarsKindCheckBox.Checked == true)
            {
                shipToDateTextBox.Text = "0.005";
            }
            else
            {
                shipToDateTextBox.Text = "";
            }

        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            nTDollarsTextBox.Text = comboBox3.Text;
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_InvoiceDSZ;
            DataRow newCustomersRow = dt2.NewRow();

            int i = rma_InvoiceDSZDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["Seqno"] = "100";
            newCustomersRow["INDescription"] = drw["INDescription"];
            newCustomersRow["InQty"] = drw["InQty"];
            newCustomersRow["UnitPrice"] = drw["UnitPrice"];
            newCustomersRow["Amount"] = drw["Amount"];
            newCustomersRow["MarkNos"] = drw["MarkNos"];
            newCustomersRow["InvoiceNo_seq"] = drw["InvoiceNo_seq"];
            newCustomersRow["Grade"] = drw["Grade"];
            newCustomersRow["InQty"] = drw["InQty"];
            newCustomersRow["size"] = drw["size"];
            newCustomersRow["RmaNo"] = drw["RmaNo"];
            newCustomersRow["VenderNo"] = drw["VenderNo"];

            newCustomersRow["CodeName"] = drw["CodeName"];
            newCustomersRow["Grade"] = drw["Grade"];
            newCustomersRow["VENDER"] = drw["VENDER"];


            try
            {
                dt2.Rows.InsertAt(newCustomersRow, rma_InvoiceDSZDataGridView.Rows.Count);
                rma_InvoiceDSZBindingSource.DataSource = dt2;


                for (int j = 0; j <= rma_InvoiceDSZDataGridView.Rows.Count - 2; j++)
                {
                    rma_InvoiceDSZDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }

                this.rma_InvoiceDSZBindingSource.EndEdit();
                this.rma_InvoiceDSZTableAdapter.Update(rm.Rma_InvoiceDSZ);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void 插入列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_InvoiceDSZ;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["SeqNo"] = 100;
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, rma_InvoiceDSZDataGridView.CurrentRow.Index);
                rma_InvoiceDSZBindingSource.DataSource = dt2;

                for (int j = 0; j <= rma_InvoiceDSZDataGridView.Rows.Count - 2; j++)
                {
                    rma_InvoiceDSZDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }

                this.rma_InvoiceDSZBindingSource.EndEdit();
                this.rma_InvoiceDSZTableAdapter.Update(rm.Rma_InvoiceDSZ);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_PackingListDSZ;
            DataRow newCustomersRow = dt2.NewRow();

            int i = rma_PackingListDSZDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["Seqno"] = "100";
            newCustomersRow["PackageNo"] = drw["PackageNo"];
            newCustomersRow["CNo"] = drw["CNo"];
            newCustomersRow["DescGoods"] = drw["DescGoods"];
            newCustomersRow["Quantity"] = drw["Quantity"];
            newCustomersRow["Net"] = drw["Net"];
            newCustomersRow["Gross"] = drw["Gross"];
            newCustomersRow["MeasurmentCM"] = drw["MeasurmentCM"];
            newCustomersRow["model"] = drw["model"];
            newCustomersRow["Ver"] = drw["Ver"];
            newCustomersRow["CardName"] = drw["CardName"];
            newCustomersRow["RmaNo"] = drw["RmaNo"];
            newCustomersRow["VenderNo"] = drw["VenderNo"];



            try
            {
                dt2.Rows.InsertAt(newCustomersRow, rma_PackingListDSZDataGridView.Rows.Count);
                rma_PackingListDSZBindingSource.DataSource = dt2;


                for (int j = 0; j <= rma_PackingListDSZDataGridView.Rows.Count - 2; j++)
                {
                    rma_PackingListDSZDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }

                this.rma_PackingListDSZBindingSource.EndEdit();
                this.rma_PackingListDSZTableAdapter.Update(rm.Rma_PackingListDSZ);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_PackingListDSZ;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["SeqNo"] = 100;
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, rma_PackingListDSZDataGridView.CurrentRow.Index);
                rma_PackingListDSZBindingSource.DataSource = dt2;

                for (int j = 0; j <= rma_PackingListDSZDataGridView.Rows.Count - 2; j++)
                {
                    rma_PackingListDSZDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }

                this.rma_PackingListDSZBindingSource.EndEdit();
                this.rma_PackingListDSZTableAdapter.Update(rm.Rma_PackingListDSZ);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void rma_PackingListDSZDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!checkBox5.Checked)
            {
                try
                {
                    if (rma_PackingListDSZDataGridView.Columns[e.ColumnIndex].Name == "QTY1" ||
                       rma_PackingListDSZDataGridView.Columns[e.ColumnIndex].Name == "Quantity")
                    {

                        int iQuantity1 = Convert.ToInt32(this.rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["QTY1"].Value);
                        int iQuantity2 = Convert.ToInt32(this.rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["Quantity"].Value);
                        this.rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["NQTY"].Value = (iQuantity1 - iQuantity2).ToString();

                    }

                    if (rma_PackingListDSZDataGridView.Columns[e.ColumnIndex].Name == "Quantity" ||
       rma_PackingListDSZDataGridView.Columns[e.ColumnIndex].Name == "Net")
                    {
                        string MODEL = rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["model2"].Value.ToString();
                        string VER = rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["Ver2"].Value.ToString();
                        string Description = rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["DescGoods"].Value.ToString().ToUpper();
                        string Quantity = rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["Quantity"].Value.ToString();
                        string Net = rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["Net"].Value.ToString();

                        if (Description.IndexOf("OPEN") != -1 && Description.IndexOf("CELL") != -1)
                        {

                            MODEL = "O" + MODEL;
                        }

                        System.Data.DataTable J1 = GetNET(MODEL, VER);
                        string NET = "";
                        if (J1.Rows.Count > 0)
                        {
                            decimal H1 = 0;
                            string H2 = J1.Rows[0][0].ToString();
                            decimal QTY2 = 0;
                            string QTY = Quantity.Replace("@", "");
                            decimal number3 = 0;
                            int number4 = 0;
                            bool canConvert = int.TryParse(QTY, out number4);
                            bool canConvert2 = decimal.TryParse(H2, out number3);
                            if (Description.IndexOf("KIT") == -1 && Description.IndexOf("TCON") == -1)
                            {
                                if (canConvert == true)
                                {
                                    QTY2 = Convert.ToDecimal(QTY);

                                    if (canConvert2 == true)
                                    {
                                        H1 = Convert.ToDecimal(H2);
                                        NET = (H1 * QTY2).ToString();
                                        this.rma_PackingListDSZDataGridView.Rows[e.RowIndex].Cells["Net"].Value = (H1 * QTY2).ToString().Replace(".0", "");
                                    }
                                }
                            }
                        }

                    }

                }
                catch { }
            }
        }
        private System.Data.DataTable GetNET(string MODEL, string VER)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CAST(ROUND(ISNULL(CASE ISNULL(CT_NW,0) WHEN '' THEN 0 ELSE CAST(RTRIM(CT_NW) AS DECIMAL(10,2)) END,0)");
            sb.Append(" /ISNULL(CASE ISNULL(CT_QTY,0) WHEN '' THEN 0 ELSE CAST(RTRIM(CT_QTY) AS DECIMAL(10,2)) END,0),1) AS DECIMAL(10,1)) WEIGHT    FROM CART");
            sb.Append(" WHERE ISNULL(CASE ISNULL(CT_NW,0) WHEN '' THEN 0 ELSE CAST(RTRIM(CT_NW) AS DECIMAL(10,2)) END,0) <> 0.00");
            sb.Append(" AND ISNULL(CASE ISNULL(CT_QTY,0) WHEN '' THEN 0 ELSE CAST(RTRIM(CT_QTY) AS DECIMAL(10,2)) END,0) <> 0.00");
            sb.Append(" AND MODEL_NO=@MODEL AND MODEL_VER=@VER");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@VER", VER));
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


        private void button8_Click(object sender, EventArgs e)
        {
            string strCollected = string.Empty;
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    if (strCollected == string.Empty)
                    {
                        strCollected = checkedListBox1.GetItemText(
         checkedListBox1.Items[i]);
                    }
                    else
                    {
                        strCollected = strCollected + checkedListBox1.
         GetItemText(checkedListBox1.Items[i]);
                    }
                }
            }
            string FD = strCollected;
            if (FD == "")
            {
                MessageBox.Show("請選澤公司");
                return;
            }
            string P = "Y";
            string USER = fmLogin.LoginID.ToString().ToUpper() + ".JPG";
            string B2 = "//acmew08r2ap//table//SIGN//USER//";
            DELETEFILE();
            string GlobalMailContent = "";
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            System.Data.DataTable T1 = GetOrderData3F(FD);

            if (T1.Rows.Count > 0)
            {
                if (checkBox1.Checked)
                {
                    string FileName2 = lsAppDir + "\\Excel\\RMA\\PACKSZ.xls";
                    if (textBox1.Text == "" || textBox2.Text == "")
                    {
                        OrderData = GetOrderData31(shippingCodeTextBox.Text);
                       
                        GetExcelProduct3(FileName2, "Y", shippingCodeTextBox.Text, boatCompanyTextBox.Text, B2 + USER, P);
                    }
                    else
                    {
                        /*
                        for (int i = 0; i <= T1.Rows.Count - 1; i++)
                        {

                            string SHIPPINGCODE = T1.Rows[i]["SHIPPINGCODE"].ToString();
                            string CARDNAME = T1.Rows[i]["CARDNAME"].ToString();
                            string COMPANY = T1.Rows[i]["COMPANY"].ToString();

                            OrderData = GetOrderData31MULTI(SHIPPINGCODE, COMPANY);
                            if (OrderData.Rows.Count > 0)
                            {
                                GetExcelProduct3(FileName2, "N", SHIPPINGCODE, CARDNAME, B2 + USER, P);
                            } if (USER == "NESSCHOU") 
                        {
                            USER = "ERINCHOU";
                        }
                        }*/
                        if (USER == "NESSCHOU.JPG")
                        {
                            USER = "ERINCHOU.JPG";
                        }
                        GetExcelProduct4(FileName2, "N", T1,  B2 + USER, P);

                    }
                }

                if (checkBox2.Checked)
                {


                    if (textBox1.Text == "" || textBox2.Text == "")
                    {
                        string FileName = lsAppDir + "\\Excel\\RMA\\還貨客戶嘜頭.xls";
                        if (GetOrderData3CNO2(shippingCodeTextBox.Text).Rows.Count == 0)
                        {
                            MessageBox.Show("請輸入CNo");
                            return;
                        }

                        OrderData = DT(shippingCodeTextBox.Text, "", "");
                        if (OrderData.Rows.Count > 0)
                        {
                            GetExcelProduct(FileName, "Y", shippingCodeTextBox.Text, "1", boatCompanyTextBox.Text);
                        }
                    }
                    else
                    {
                        string FileName = lsAppDir + "\\Excel\\RMA\\還貨客戶嘜頭2.xls";
                        System.Data.DataTable T1C = GetOrderData3FCREATE();
                        for (int i = 0; i <= T1C.Rows.Count - 1; i++)
                        {

                            string CREATEDATE = T1C.Rows[i]["CREATEDATE"].ToString();
                            string COMPANY = T1C.Rows[i]["COMPANY"].ToString();

                            OrderData = DT("", CREATEDATE, COMPANY);
                            if (OrderData.Rows.Count > 0)
                            {
                                GetExcelProductCREATE(FileName, "N", CREATEDATE);
                            }
                        }

                    }
                }

                if (checkBox3.Checked)
                {
                    string ExcelTemplate = lsAppDir + "\\Excel\\RMA\\AU自提簽收單.xls";
                    string OutPutFile = "";
                    if (textBox1.Text == "" || textBox2.Text == "")
                    {
                        System.Data.DataTable OrderData = DT(shippingCodeTextBox.Text, "", "");
                        OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\" +
"AUO_2nd RMA 自提工單.xls";
                        if (OrderData.Rows.Count > 0)
                        {
                            ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "Y");
                        }
                    }
                    else
                    {
                        for (int i = 0; i <= T1.Rows.Count - 1; i++)
                        {
                            string SHIPPINGCODE = T1.Rows[i]["SHIPPINGCODE"].ToString();
                            string COMPANY = T1.Rows[i]["COMPANY"].ToString();

                            System.Data.DataTable OrderData = DT(SHIPPINGCODE, "", COMPANY);

                            if (OrderData.Rows.Count > 0)
                            {
                                OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\" +
            "AUO_2nd RMA 自提工單" + " " + (i + 1).ToString() + ".xls";
                                ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");

                                SetMsg("AUO_2nd RMA 自提工單" + " " + (i + 1).ToString() + ".xls");
                            }

                        }
                    }
                }
                if (textBox1.Text == "" || textBox2.Text == "")
                {
                }
                else
                {
                    if (checkBox4.Checked)
                    {
                        string ExcelTemplate = lsAppDir + "\\Excel\\RMA\\SZ還貨打包明細總表.xls";
                        string OutPutFile = "";
                        System.Data.DataTable H1 = DTG();
                        dataGridView1.DataSource = H1;
                        GlobalMailContent = htmlMessageBody(dataGridView1).ToString();

                        System.Data.DataTable OrderData = GetOrderData3GS(FD);
                        OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\" +
"SZ還貨打包明細總表.xls";
                        if (OrderData.Rows.Count > 0)
                        {
                            ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");
                        }

                    }
                    else
                    {
                        GlobalMailContent = "";
                    }
                    string SUBJECT = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_归还客户RMA清单";
                    string DATA = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_归还客户RMA 资料清单! 请参考～";
                    string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";


                    ExcelReport.MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, GlobalMailContent, DATA);
                }
                SetMsg("匯出成功");

            }
            else
            {
                SetMsg("沒有資料");
            }
        }

        private void SetMsg(string Msg)
        {
            lblMsg.Text = "處理訊息:" + Msg;
            lblMsg.Refresh();
        }


        private StringBuilder htmlMessageBody(DataGridView dg)
        {

            string KeyValue = "";

            string tmpKeyValue = "";

            StringBuilder strB = new StringBuilder();

            if (dg.Rows.Count == 0)
            {
                strB.AppendLine("<table class='GridBorder' cellspacing='0'");
                strB.AppendLine("<tr><td>***  今日無資料  ***</td></tr>");
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

                    //if (i != 0)
                    //{
                    //    strB.AppendLine("<tr class='HeaderBorder'>");
                    //    for (int iCol = 0; iCol < dg.Columns.Count; iCol++)
                    //    {
                    //        strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
                    //    }
                    //    strB.AppendLine("</tr>");
                    //}

                    //處理鍵值




                    KeyValue = dg.Rows[i].Cells[0].Value.ToString();
                    tmpKeyValue = KeyValue;

                    if (KeyValue.IndexOf("{") >= 0)
                    {
                        tmpKeyValue = "";
                    }
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


                DataGridViewCell dgvc;


                if (string.IsNullOrEmpty(tmpKeyValue))
                {
                    strB.AppendLine("<td>&nbsp;</td>");
                }
                else
                {
                    strB.AppendLine("<td align='center'>" + tmpKeyValue + "</td>");
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
                            strB.AppendLine("<td align='right'>" + x.ToString("#,##0") + "</td>");
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
                                    string sDate = "";
                                    if (dgvc.Value.ToString() != "小計")
                                    {
                                        sDate = dgvc.Value.ToString().Substring(0, 4) + "/" +
                                                    dgvc.Value.ToString().Substring(4, 2) + "/" +
                                                    dgvc.Value.ToString().Substring(6, 2);
                                        strB.AppendLine("<td>" + sDate + "</td>");
                                    }
                                    else
                                    {
                                        strB.AppendLine("<td>" + dgvc.Value.ToString() + "</td>");
                                    }


                                }
                            }

                            else
                            {
                                strB.AppendLine("<td align='center'>" + dgvc.Value.ToString() + "</td>");
                            }
                        }

                    }


                }
                strB.AppendLine("</tr>");

            }
            //table footer & end of html file
            //strB.AppendLine("</table></center></body></html>");
            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }







    }
}


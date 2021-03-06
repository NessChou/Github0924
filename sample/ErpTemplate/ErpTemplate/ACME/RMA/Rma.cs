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
    public partial class Rma : ACME.fmBase1
    {
        int CON = 0;
        string H1 = "";
        string NewFileName = "";
        private System.Data.DataTable OrderData;
        public Rma()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            rma_mainTableAdapter.Connection = MyConnection;
            rma_InvoiceDTableAdapter.Connection = MyConnection;
            rma_InvoiceD2TableAdapter.Connection = MyConnection;
            rma_PackingListDTableAdapter.Connection = MyConnection;
            rma_MarkTableAdapter.Connection = MyConnection;
            rma_Mark2TableAdapter.Connection = MyConnection;
            rMA_LADINGDTableAdapter.Connection = MyConnection;
            dataTable2TableAdapter.Connection = MyConnection;
        }

        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
        }
 
   
        public override void AfterEdit()
        {


            modifyNameTextBox.Text = fmLogin.LoginID.ToString();
            shippingCodeTextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
     
        }
        public override void AfterCancelEdit()
        {
            Control();

        }
        private void Control()
        {

            createNameTextBox.ReadOnly = true;
            modifyNameTextBox.ReadOnly = true;
            shippingCodeTextBox.ReadOnly = true;

            button3.Enabled = true;
            button7.Enabled = true;
            btnEmailPeter.Enabled = true;
            button8.Enabled = true;
            button10.Enabled = true;
            button13.Enabled = true;
            button15.Enabled = true;
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
                rm.Rma_InvoiceD.RejectChanges();
                rm.Rma_InvoiceD2.RejectChanges();
                rm.Rma_PackingListD.RejectChanges();
                rm.Rma_Mark.RejectChanges();
                rm.Rma_Mark2.RejectChanges();
                rm.RMA_LADINGD.RejectChanges();

            }
            catch
            {
            }

            return true;
        }
        private void CalcTotals2()
        {
            try
            {

                Int32 Quantity = 0;
                decimal NET = 0;
                decimal GROSS = 0;


                int i = this.rma_PackingListDDataGridView.Rows.Count - 2;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {
                    if (!String.IsNullOrEmpty(rma_PackingListDDataGridView.Rows[iRecs].Cells["Quantity"].Value.ToString()))
                    {
                        int g = rma_PackingListDDataGridView.Rows[iRecs].Cells["Quantity"].Value.ToString().LastIndexOf("@");
                        if (g != 0)
                        {
                            Quantity += Convert.ToInt32(rma_PackingListDDataGridView.Rows[iRecs].Cells["Quantity"].Value);

                        }
                    }
                    if (!String.IsNullOrEmpty(rma_PackingListDDataGridView.Rows[iRecs].Cells["Net"].Value.ToString()))
                    {
                        int U = rma_PackingListDDataGridView.Rows[iRecs].Cells["Net"].Value.ToString().LastIndexOf("@");
                        if (U != 0)
                        {

                            NET += Convert.ToDecimal(rma_PackingListDDataGridView.Rows[iRecs].Cells["Net"].Value);
                        }
                    }

                    if (!String.IsNullOrEmpty(rma_PackingListDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn56"].Value.ToString()))
                    {

                        int V = rma_PackingListDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn56"].Value.ToString().LastIndexOf("@");
                        if (V != 0)
                        {
                            GROSS += Convert.ToDecimal(rma_PackingListDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn56"].Value);
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
        private void CalcTotals1()
        {
            try
            {

                Int32 AMT = 0;



                int i = this.rma_InvoiceDDataGridView.Rows.Count - 2;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {

                    if (!String.IsNullOrEmpty(rma_InvoiceDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn45"].Value.ToString()))
                    {
                        AMT += Convert.ToInt32(rma_InvoiceDDataGridView.Rows[iRecs].Cells["dataGridViewTextBoxColumn45"].Value);
                    }
                    
                    
                   
                }

                if (AMT != 0)
                {
                    add8TextBox.Text = new Class1().NumberToString(AMT);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        public override void AfterEndEdit()
        {
            try
            {

    
                    if (rma_InvoiceDDataGridView.Rows.Count > 1)
                    {

                        CalcTotals1();

                        System.Data.DataTable dt1 = rm.Rma_InvoiceD;
                        try
                        {
                            string D1 = quantityTextBox.Text;
                            if (D1 != "")
                            {
                                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                                {
                                    DateTime R1 = Convert.ToDateTime(GetMenu.DayS(D1));

                                    DataRow drw = dt1.Rows[i];
                                    string aa = drw["RmaNo"].ToString();
                                    string bb = drw["shippingcode"].ToString();
                                    string InQty = drw["InQty"].ToString();
                                    UPDATEJOBNO(bb, R1, InQty, aa);
                                }
                            }
                        }
                        catch { }

                        try
                        {
                            string D2 = closeDayTextBox.Text;
                            if (D2 != "")
                            {
                                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                                {
                                    DateTime R1 = Convert.ToDateTime(GetMenu.DayS(D2));

                                    DataRow drw = dt1.Rows[i];
                                    string aa = drw["RmaNo"].ToString();
                                    string bb = drw["shippingcode"].ToString();
                                    string InQty = drw["InQty"].ToString();
                                    UPDATEJOBNO2(bb, R1, InQty, aa);
                                }
                            }
                        }
                        catch { }


                        try
                        {
                            string D3 = arriveDayTextBox.Text;
                            if (D3 != "")
                            {
                                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                                {
                                    DateTime R1 = Convert.ToDateTime(GetMenu.DayS(D3));

                                    DataRow drw = dt1.Rows[i];
                                    string aa = drw["RmaNo"].ToString();
                                    string bb = drw["shippingcode"].ToString();
                                    string InQty = drw["InQty"].ToString();
                                    UPDATEJOBNO3(bb, R1, InQty, aa);
                                }
                            }
                        }
                        catch { }
                    }
               
                    if (rma_PackingListDDataGridView.Rows.Count > 1)
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

                    }

                
                rma_mainBindingSource.EndEdit();
                rma_mainTableAdapter.Update(rm.Rma_main);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        public void UPDATEJOBNO(string u_jobno,DateTime U_ACME_BackDate1,string u_acme_backqty1, string u_rma_no)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("update octr set u_jobno=@u_jobno,U_ACME_BackDate1=@U_ACME_BackDate1,u_acme_backqty1=@u_acme_backqty1 where u_rma_no=@u_rma_no", connection);
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
        public void UPDATEJOBNO2(string u_jobno, DateTime U_ACME_Out, string U_yetqty, string u_rma_no)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("update octr set u_jobno=@u_jobno,U_ACME_Out=@U_ACME_Out,U_yetqty=@U_yetqty where u_rma_no=@u_rma_no", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@u_jobno", u_jobno));
            command.Parameters.Add(new SqlParameter("@U_ACME_Out", U_ACME_Out));
            command.Parameters.Add(new SqlParameter("@U_yetqty", U_yetqty));
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
        public void UPDATEJOBNO3(string u_jobno, DateTime U_ACME_BackDate, string U_ACME_QBack, string u_rma_no)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("update octr set u_jobno=@u_jobno,U_ACME_BackDate=@U_ACME_BackDate,U_ACME_QBack=@U_ACME_QBack where u_rma_no=@u_rma_no", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@u_jobno", u_jobno));
            command.Parameters.Add(new SqlParameter("@U_ACME_BackDate", U_ACME_BackDate));
            command.Parameters.Add(new SqlParameter("@U_ACME_QBack", U_ACME_QBack));
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
        public override void SetInit()
        {

            MyBS = rma_mainBindingSource;
            MyTableName = "Rma_main";
            MyIDFieldName = "ShippingCode";
             UtilSimple.SetLookupBinding(receiveDayComboBox, "receiveDay", rma_mainBindingSource, "receiveDay");
             UtilSimple.SetLookupBinding(boardCountNoComboBox, "boardCountNo", rma_mainBindingSource, "boardCountNo");
             UtilSimple.SetLookupBinding(shipToDateComboBox, "shipToDate", rma_mainBindingSource, "shipToDate");

             //處理複製
             MasterTable = rm.Rma_main;
             DetailTables = new System.Data.DataTable[] { rm.Rma_InvoiceD };
             DetailBindingSources = new BindingSource[] { rma_InvoiceDBindingSource1 };

         }
         public override void SetDefaultValue()
         {
             if (kyes == null)
             {

                 string NumberName = "RMA" + DateTime.Now.ToString("yyyyMMdd");
                 string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                 kyes = NumberName + AutoNum + "X";
             }
             this.shippingCodeTextBox.Text = kyes;
             kyes = this.shippingCodeTextBox.Text;
             createNameTextBox.Text = fmLogin.LoginID.ToString();
             cFSCheckBox.Checked = false;
             boatNameCheckBox.Checked = false;
             this.rma_mainBindingSource.EndEdit();
             kyes = null;
         }
         public override void AfterCopy()
         {
             if (kyes == null)
             {
                 string NumberName = "RMA" + DateTime.Now.ToString("yyyyMMdd");
                 string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                 this.shippingCodeTextBox.Text = NumberName + AutoNum + "X";
                 kyes = this.shippingCodeTextBox.Text;
             }
         }
        public override void FillData()
        {
            try
            {
                //if (shippingCodeTextBox.Text != "")
                //{
                //    MyID = shippingCodeTextBox.Text;
                //}
                dataTable2TableAdapter.Fill(rm.DataTable2, MyID);
                rma_mainTableAdapter.Fill(rm.Rma_main, MyID);      
                rma_InvoiceDTableAdapter.Fill(rm.Rma_InvoiceD, MyID);
                rma_InvoiceD2TableAdapter.Fill(rm.Rma_InvoiceD2, MyID);
                rma_DownloadTableAdapter.Fill(rm.Rma_Download, MyID);
                rma_PackingListDTableAdapter.Fill(rm.Rma_PackingListD, MyID);
                rma_MarkTableAdapter.Fill(rm.Rma_Mark, MyID);
                rma_Mark2TableAdapter.Fill(rm.Rma_Mark2, MyID);

                rMA_LADINGDTableAdapter.Fill(rm.RMA_LADINGD, MyID);

                checkBox1.Checked = false;
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

                rma_InvoiceDBindingSource1.MoveFirst();
                for (int i = 0; i <= rma_InvoiceDBindingSource1.Count - 1; i++)
                {
                    DataRowView row = (DataRowView)rma_InvoiceDBindingSource1.Current;
                    row["SeqNo"] = i;
                    rma_InvoiceDBindingSource1.EndEdit();
                    rma_InvoiceDBindingSource1.MoveNext();
                }

                rma_InvoiceD2BindingSource.MoveFirst();
                for (int i = 0; i <= rma_InvoiceD2BindingSource.Count - 1; i++)
                {
                    DataRowView row = (DataRowView)rma_InvoiceD2BindingSource.Current;
                    row["SeqNo"] = i;
                    rma_InvoiceD2BindingSource.EndEdit();
                    rma_InvoiceD2BindingSource.MoveNext();
                }

                rma_MarkBindingSource.MoveFirst();
                for (int i = 0; i <= rma_MarkBindingSource.Count - 1; i++)
                {
                    DataRowView row1 = (DataRowView)rma_MarkBindingSource.Current;
                    row1["Seq"] = i;
                    rma_MarkBindingSource.EndEdit();
                    rma_MarkBindingSource.MoveNext();
                }

            


                rma_PackingListDBindingSource.MoveFirst();
                for (int i = 0; i <= rma_PackingListDBindingSource.Count - 1; i++)
                {
                    DataRowView row2 = (DataRowView)rma_PackingListDBindingSource.Current;
                    row2["SeqNo"] = i;
                    rma_PackingListDBindingSource.EndEdit();
                    rma_PackingListDBindingSource.MoveNext();
                }

                rMA_LADINGDBindingSource.MoveFirst();
                for (int i = 0; i <= rMA_LADINGDBindingSource.Count - 1; i++)
                {
                    DataRowView row3 = (DataRowView)rMA_LADINGDBindingSource.Current;
                    row3["SeqNo"] = i;
                    rMA_LADINGDBindingSource.EndEdit();
                    rMA_LADINGDBindingSource.MoveNext();
                }

                rma_mainTableAdapter.Connection.Open();

              
                rma_mainBindingSource.EndEdit();
                rma_InvoiceDBindingSource1.EndEdit();
                rma_InvoiceD2BindingSource.EndEdit();
                rma_DownloadBindingSource.EndEdit();
                rma_PackingListDBindingSource.EndEdit();
                rma_MarkBindingSource.EndEdit();
                rma_Mark2BindingSource.EndEdit();
                rMA_LADINGDBindingSource.EndEdit();
  
                rma_mainTableAdapter.Update(rm.Rma_main);
                rMA_LADINGDTableAdapter.Update(rm.RMA_LADINGD);
                rma_InvoiceDTableAdapter.Update(rm.Rma_InvoiceD);
                rma_InvoiceD2TableAdapter.Update(rm.Rma_InvoiceD2);
                rma_DownloadTableAdapter.Update(rm.Rma_Download);
                rma_PackingListDTableAdapter.Update(rm.Rma_PackingListD);
                rma_MarkTableAdapter.Update(rm.Rma_Mark);
                rma_Mark2TableAdapter.Update(rm.Rma_Mark2);
   
               
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
                this.rma_mainTableAdapter.Connection.Close();

            }
            return UpdateData;
        }



        private void button2_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuList();

            if (LookupValues != null)
            {
              
                cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                cardNameTextBox.Text = Convert.ToString(LookupValues[1]);
                add9TextBox.Text = Convert.ToString(LookupValues[5]);
                add10TextBox.Text = Convert.ToString(LookupValues[6]);
                boatCompanyTextBox.Text = Convert.ToString(LookupValues[2]);
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
            string P1 = "";
            frm1.q1 = H1;

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (radioButton1.Checked)
                    {
                        P1 = frm1.q;
                    }
                    if (radioButton2.Checked)
                    {
                        P1 = frm1.q2;
                    }
                    System.Data.DataTable dt1 = GetAR2(H1,P1);
                    System.Data.DataTable dt2 = rm.Rma_InvoiceD;
                
                        for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                        {
                            DataRow drw = dt1.Rows[i];
                            DataRow drw2 = dt2.NewRow();
                            drw2["ShippingCode"] = shippingCodeTextBox.Text;
                            drw2["RmaNo"] = drw["U_RMA_NO"];
                            drw2["MarkNos"] = drw["U_RMODEL"];
                            drw2["SeqNo"] = "0";
                            drw2["VenderNo"] = drw["U_AUO_RMA_NO"];

                            string SIZE = drw["bb"].ToString();

                            drw2["size"] = SIZE;
                            string T1 = drw["aa"].ToString();
                            string MODEL = drw["U_RMODEL"].ToString();
                     
                            string SIZE2 = "";
                            int I2 = MODEL.ToUpper().IndexOf("_OPEN CELL");
                            int I3 = MODEL.ToUpper().IndexOf("KIT_");
                            int I4 = MODEL.ToUpper().IndexOf("KIT AD BOARD");
                            int I5 = MODEL.ToUpper().IndexOf("KIT AD INVERTER");
                            int I6 = MODEL.ToUpper().IndexOf("KIT DRIVER BOARD");
                            int I7 = MODEL.ToUpper().IndexOf("KIT INVERTER");
                            int I8 = MODEL.ToUpper().IndexOf("_T CON");
                            int I12 = MODEL.ToLower().IndexOf("enclosure");
                             
                              //OpenFrame 
                            int I9 = MODEL.ToUpper().IndexOf("\"");
                            int I11 = MODEL.ToUpper().IndexOf("”");
                            int I10 = MODEL.ToUpper().IndexOf("OPENFRAME");
                              if (I9 != -1)
                              {
                                  SIZE2 = MODEL.Substring(0, I9);
                              }
                              if (I11 != -1)
                              {
                                  SIZE2 = MODEL.Substring(0, I11);
                              }


                              if (I2 != -1)
                            {
                                  
                             //   drw2["MarkNos"] = drw["U_RMODEL"].ToString().Substring(0, I2);
                                drw2["INDescription"] = SIZE + "\" OPEN CELL_AU";
             
                                  
                            }
                              else if (I8 != -1)
                              {
                                //  drw2["MarkNos"] = drw["U_RMODEL"].ToString().Substring(0, I8);
                                  drw2["INDescription"] = "TCON-PCBA_";
                 
                              }
                              else if (I3 != -1)
                              {
                                //  drw2["MarkNos"] = SIZE2 + "\" KIT";
                                  drw2["INDescription"] = SIZE2 + "\" KIT";
                                  drw2["DIFF"] = "Y";
                              }
                              else if (I4 != -1)
                              {
                               //   drw2["MarkNos"] = SIZE2+ "\" KIT_AD Board";
                                  drw2["INDescription"] = SIZE2 + "\" KIT_AD Board";
                                  drw2["DIFF"] = "Y";
                              }
                              else if (I5 != -1)
                              {
                                //  drw2["MarkNos"] = SIZE2 + "\" KIT_Inverter";
                                  drw2["INDescription"] = SIZE2 + "\" KIT_Inverter";
                                  drw2["DIFF"] = "Y";
                              }
                              else if (I6 != -1)
                              {
                                  //drw2["MarkNos"] = SIZE2 + "\" KIT_Driver Board";
                                  drw2["INDescription"] = SIZE2 + "\" KIT_Driver Board";
                                  drw2["DIFF"] = "Y";
                              }
                              else if (I7 != -1)
                              {
                                //  drw2["MarkNos"] = SIZE2 + "\" KIT_Inverter";
                                  drw2["INDescription"] = SIZE2 + "\" KIT_Inverter";
                                  drw2["DIFF"] = "Y";
                              }
                              else if (I10 != -1)
                              {
                   
                                //  drw2["MarkNos"] =  SIZE2 + "\" OpenFrame";
                                  drw2["INDescription"] = "OFD_" + SIZE2 + "\" OpenFrame LCD Monitor";
                                  drw2["DIFF"] = "Y";
                              }
                              else if (I12 != -1)
                              {

                                //  drw2["MarkNos"] = SIZE2 + "\" Enclosure Display";
                                  drw2["INDescription"] =  SIZE2 + "\" Enclosure Display";
                                  drw2["DIFF"] = "Y";
                              }
                              else
                              {

                  
                               //   drw2["MarkNos"] = MODEL;
                                  System.Data.DataTable J1 = GetDESC(MODEL);
                                  if (J1.Rows.Count > 0)
                                  {
                                      drw2["INDescription"] = J1.Rows[0][0].ToString() ;
                                  }
                                  else
                                  {
                                      drw2["INDescription"] = T1;
                                  }
                              }
                         
                            
                            int g = drw["U_Rquinity"].ToString().IndexOf("+");
                            int t = drw["U_Rquinity"].ToString().LastIndexOf("+");
                            string h;
                            string s;
              
                  
                            if (g == -1)
                            {
                                s = drw["U_Rquinity"].ToString();
                                drw2["InQty"] = s;
                            }
                            else
                            {
                                s = drw["U_Rquinity"].ToString().Substring(g + 1);

                                try
                                {

                                    if (drw["U_Rquinity"].ToString().Substring(3, 1) != "+")
                                    {
                                        h = drw["U_Rquinity"].ToString().Substring(0, 2);
                                    }
                                    else
                                    {
                                        h = drw["U_Rquinity"].ToString().Substring(0, 1);
                                    }
                                    int a = Convert.ToInt16(s.ToString());
                                    int b = Convert.ToInt16(h.ToString());
                                    drw2["InQty"] = (a + b).ToString();
                                }
                                catch (Exception ex)
                                {
                                    h = drw["U_Rquinity"].ToString().Substring(0, 1);
                                    drw2["InQty"] = h.ToString();
                                }
                     
                            }
                            string GRADE = drw["U_Rgrade"].ToString();
                            string VER = drw["U_Rver"].ToString();

                            drw2["Grade"] = GRADE;

                            drw2["InvoiceNo_seq"] = VER;
            
                            drw2["CodeName"] = drw["U_cusname_s"];
                            drw2["VENDER"] = drw["U_Rvender"];
                            string CARD = cardCodeTextBox.Text.Trim();
                            if (CARD == "S0001-00")
                            {
                                CARD = "S0001-GD";
                            }
                            System.Data.DataTable G1 = GetOCRD(CARD);
                            if (G1.Rows.Count > 0)
                            {
                                string CTYPE = G1.Rows[0][0].ToString();
                                System.Data.DataTable G2 = null;
                                GRADE = GRADE.Substring(0, 1);
                                if (GRADE == "/")
                                {
                                    GRADE = "Z";
                                }
                                if (CTYPE == "C")
                                {
                                    G2 = GetOCRD2(CARD, MODEL, GRADE, VER);
                                }
                                else
                                {
                                    string DOCTYPE = "";
                                    if (CARD == "S0001-GD")
                                    {
                                        DOCTYPE = "1";
                                    }
                                    G2 = GetOCRD3(CARD, MODEL, GRADE, VER, DOCTYPE);
                                }
                                if (G2.Rows.Count > 0)
                                {
                                    drw2["INV"] = G2.Rows[0]["INV"].ToString();
                                    drw2["PRICE"] = Convert.ToDecimal(G2.Rows[0]["PRICE"]);
                                }
                            }
                            dt2.Rows.Add(drw2);
                        }

                        for (int j = 0; j <= rma_InvoiceDDataGridView.Rows.Count - 2; j++)
                        {
                            rma_InvoiceDDataGridView.Rows[j].Cells[0].Value = j.ToString();
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
        public static System.Data.DataTable GetAR2(string COMPANY,string DocEntry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string aa = '"'.ToString();
            string sql = "";

            if (COMPANY=="進金生")
            {
                sql = "select U_RMA_NO,U_AUO_RMA_NO,U_Rgrade,U_Rmodel,U_cusname_s,U_Rver,U_Rmodel,U_Rvender,U_Rquinity,aa=case substring(U_Rmodel,4,1) when '0' then substring(U_Rmodel,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1)+'" + aa + "'+'TFT LCD MODULE' END,bb=case substring(U_Rmodel,4,1) when '0' then substring(U_Rmodel,2,2) ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1) END from ACMESQL02.DBO.octr where Contractid IN (" + DocEntry + ") ";
            }
            if (COMPANY == "達睿生")
            {
                sql = "select U_RMA_NO,U_AUO_RMA_NO,U_Rgrade,U_Rmodel,U_cusname_s,U_Rver,U_Rmodel,U_Rvender,U_Rquinity,aa=case substring(U_Rmodel,4,1) when '0' then substring(U_Rmodel,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1)+'" + aa + "'+'TFT LCD MODULE' END,bb=case substring(U_Rmodel,4,1) when '0' then substring(U_Rmodel,2,2) ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1) END from ACMESQL05.DBO.octr where Contractid IN (" + DocEntry + ") ";
            }
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
        public static System.Data.DataTable GetOCRD(string CARDCODE)
        {
            SqlConnection MyConnection = globals.shipConnection;

            string sql = "SELECT CARDTYPE  FROM OCRD WHERE CARDCODE=@CARDCODE ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
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
        public static System.Data.DataTable GetOCRD2(string CARDCODE, string MODEL, string GRADE, string VERSION)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT TOP 1  T0.U_ACME_INV INV,CAST(ROUND(T5.PRICE,2) AS decimal(18,2)) PRICE FROM ODLN T0");
            sb.Append("   LEFT JOIN DLN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("  left join RDR1 T5 on (T1.baseentry=T5.docentry and  T1.baseline=T5.linenum  and T5.targettype='15')");
            sb.Append("  WHERE T0.CARDCODE=@CARDCODE ");
            sb.Append(" AND T1.ITEMCODE IN (SELECT ITEMCODE FROM AcmeSql02.DBO.OITM ");
            sb.Append(" WHERE U_TMODEL  =@MODEL  AND U_GRADE =@GRADE AND U_VERSION =@VERSION)");
            sb.Append(" ORDER BY T0.DOCENTRY DESC");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@VERSION", VERSION));
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
        public static System.Data.DataTable GetOCRD3(string CARDCODE, string MODEL, string GRADE, string VERSION, string DOCTYPE)
        {
            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT TOP 1  T0.U_ACME_INV INV,CAST(ROUND(T5.PRICE,2) AS decimal(18,2)) PRICE FROM OPDN T0");
            sb.Append("   LEFT JOIN PDN1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("  left join POR1 T5 on (T1.baseentry=T5.docentry and  T1.baseline=T5.linenum  and T5.targettype='20')");
            if (DOCTYPE == "1")
            {
                sb.Append("  WHERE T0.CARDCODE LIKE '%S0001%' ");
            }
            else
            {
                sb.Append("  WHERE T0.CARDCODE=@CARDCODE");
            }
            sb.Append(" AND T1.ITEMCODE IN (SELECT ITEMCODE FROM AcmeSql02.DBO.OITM ");
            sb.Append(" WHERE U_TMODEL  =@MODEL  AND U_GRADE =@GRADE AND U_VERSION =@VERSION)");
            sb.Append(" ORDER BY T0.DOCENTRY DESC");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@MODEL", MODEL));
            command.Parameters.Add(new SqlParameter("@GRADE", GRADE));
            command.Parameters.Add(new SqlParameter("@VERSION", VERSION));
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
        private void Rma_Load(object sender, EventArgs e)
        {

            Control();

         
    

            //上傳
            DataGridViewLinkColumn column = new DataGridViewLinkColumn();

            column.Name = "Link";
            column.UseColumnTextForLinkValue = true;

            column.Text = "讀取檔案";
            column.LinkBehavior = LinkBehavior.HoverUnderline;

            column.TrackVisitedState = true;

           rma_DownloadDataGridView.Columns.Add(column);
            //

           textBox2.Text = fmLogin.LoginID.ToString() + "@acmepoint.com";


           if (globals.GroupID.ToString().Trim() != "EEP")
           {

               textBox1.Visible = false;
               textBox3.Visible = false;
               textBox4.Visible = false;
               textBox5.Visible = false;

           }


           System.Data.DataTable dt4 = GetMenu.GETRMAWH();


           comboBox1.Items.Clear();


           for (int i = 0; i <= dt4.Rows.Count - 1; i++)
           {
               comboBox1.Items.Add(Convert.ToString(dt4.Rows[i][0]));
           }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult result;
            result = MessageBox.Show("收件人地址為" + textBox2.Text + "是否要寄出", "Close", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {



                string template;
                StreamReader objReader;
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\MailTemplates\\Report.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();


                StringWriter writer = new StringWriter();
                HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
                string Html = GetTODO_USERDataSource2();

                template = template.Replace("##shippingCode##", "JOB NO: " + shippingCodeTextBox.Text);
                template = template.Replace("##tradeCondition##", "貿易條件 : " + tradeConditionTextBox1.Text);
                template = template.Replace("##goalPlace##", "目的地 : " + goalPlaceTextBox.Text);
                template = template.Replace("##shipment##", "裝船港 : " + shipmentTextBox.Text);
                template = template.Replace("##unloadCargo##", "卸貨港 : " + unloadCargoTextBox.Text);
                template = template.Replace("##boardCount##", "20呎 : " + boardCountTextBox.Text);
                template = template.Replace("##boardDeliver##", "40呎 : " + boardDeliverTextBox.Text);
                template = template.Replace("##sendGoods##", "併櫃/CBM : " + sendGoodsTextBox.Text);
                template = template.Replace("##receiveDay##", "運送方式 : " + receiveDayComboBox.Text);
                template = template.Replace("##boardCountNo##", "貿易形式 : " + boardCountNoComboBox.Text);
                template = template.Replace("##Content##", Html);


                MailMessage message = new MailMessage();

                string aa = textBox2.Text;
   
                message.To.Add(new MailAddress(aa));

                message.Subject = "ShippingOrder";
                message.Body = template;
           
                //格式為 Html
                message.IsBodyHtml = true;
     
                SmtpClient client = new SmtpClient();
                try
                {
                    client.Send(message);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
            }
        }

        private string GetTODO_USERDataSource2()
        {
            System.Data.DataTable dtEvent = GetMenu.GetMail(shippingCodeTextBox.Text);

            string html = string.Empty;
            string DateGroup = string.Empty;

            foreach (DataRow row in dtEvent.Rows)
            {
                string Docentry = Convert.ToString(row["Docentry"]);
                string itemcode = Convert.ToString(row["itemcode"]);
                string Dscription = Convert.ToString(row["Dscription"]);
                string Quantity = Convert.ToString(row["Quantity"]);
                html = html + "<tr ><td>" + Docentry + "</td><td>" + itemcode + "</td><td>" + Dscription + "</td><td>" + Quantity + "</td></tr>";
            }
            return html;
        }
        public static System.Data.DataTable GetMail(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select * from dbo.Rma_Item where shippingcode=@shippingcode";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Shipping_Item");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["Shipping_Item"];
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                string server = "//acmesrv01//SAP_Share//Rma//";

            
                    OpenFileDialog opdf = new OpenFileDialog();
                    DialogResult result = opdf.ShowDialog();
                    string filename = Path.GetFileName(opdf.FileName);
                    System.Data.DataTable dt2 = GetMenu.download(filename);
                    if (dt2.Rows.Count > 0)
                    {
                        MessageBox.Show("檔案名稱重複,請修改檔名");
                    }
                    else
                    {
                        if (result == DialogResult.OK)
                        {
                            MessageBox.Show(Path.GetFileName(opdf.FileName));
                            string file = opdf.FileName;
                            bool FF1 = getrma.UploadFile(file, server, false);
                            if (FF1 == false)
                            {
                                return;
                            }
                            System.Data.DataTable dt1 = rm.Rma_Download;
                            DataRow drw = dt1.NewRow();
                            drw["ShippingCode"] = shippingCodeTextBox.Text;
                            drw["seq"] = (rma_DownloadDataGridView.Rows.Count).ToString();
                            drw["filename"] = filename;
                            drw["path"] = @"\\acmesrv01\SAP_Share\Rma\" + filename;
                            dt1.Rows.Add(drw);
                            this.rma_DownloadBindingSource.EndEdit();
                            this.rma_DownloadTableAdapter.Update(rm.Rma_Download);
                        }
                    }
    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void rma_DownloadDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "Link")
                {
                    System.Data.DataTable dt1 = rm.Rma_Download;
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

      

    


        private System.Data.DataTable GetOrderData4()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select top 1 packageno,seqno,cno from Rma_PackingListD");
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



        public static System.Data.DataTable rmainvoice(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from rma_invoicem where shippingcode=@shippingcode ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
 

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rma_invoicem");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rma_invoicem"];
        }

        public static System.Data.DataTable rmapcak(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from Rma_PackingListM where shippingcode=@shippingcode ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "Rma_PackingListM");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["Rma_PackingListM"];
        }

       

      
        public static System.Data.DataTable Getshipitem(string shippingcode)
        {
            string aa = '"'.ToString();
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select aa=case substring(dscription,4,1) when 0 then substring(dscription,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(dscription,2,2)+'.'+substring(dscription,4,1)+'" + aa + "'+'TFT LCD MODULE' END,DSCRIPTION,QUANTITY,'V.'+PINO PINO,bb=case substring(dscription,4,1) when 0 then substring(dscription,2,2) ELSE substring(dscription,2,2)+'.'+substring(dscription,4,1) END  from Rma_Item where shippingcode=@shippingcode  ");
         //   string sql = "select * from Rma_Item where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "Rma_Item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["Rma_Item"];
        }

        public static System.Data.DataTable Getshipitem2(string shippingcode,string seqno)
        {
            string aa = '"'.ToString();
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select aa=case substring(dscription,4,1) when 0 then substring(dscription,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(dscription,2,2)+'.'+substring(dscription,4,1)+'" + aa + "'+'TFT LCD MODULE' END,DSCRIPTION,QUANTITY,'V.'+PINO PINO,bb=case substring(dscription,4,1) when 0 then substring(dscription,2,2) ELSE substring(dscription,2,2)+'.'+substring(dscription,4,1) END  from Rma_Item where shippingcode=@shippingcode and seqno=@seqno  ");
            //   string sql = "select * from Rma_Item where shippingcode=@shippingcode";
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seqno", seqno));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "Rma_Item");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["Rma_Item"];
        }

    


        private void button10_Click(object sender, EventArgs e)
        {
            string P = "Y";
            string USER = fmLogin.LoginID.ToString().ToUpper() + ".JPG";
            string B2 = "//acmew08r2ap//table//SIGN//USER//";
            System.Data.DataTable dtmark1 = GetAuono(shippingCodeTextBox.Text, add9TextBox.Text);
            if (dtmark1.Rows.Count < 1 )
            {
                MessageBox.Show("請先存檔");
                return;
            }
            else if (dtmark1.Rows.Count > 13)
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                FileName = lsAppDir + "\\Excel\\RMA\\INVO2.xls";
                OrderData = GetOrderData2();
                GetExcelProduct4(FileName, B2 + USER, P);

            }
            else
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                FileName = lsAppDir + "\\Excel\\RMA\\INVO.xls";
                OrderData = GetOrderData2();
                GetExcelProduct3(FileName, B2 + USER, P);
            }
               
         

        }

        private void GetExcelProduct3(string ExcelFile,string P1,string FLAG)
        {

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
            excelSheet.Name = shippingCodeTextBox.Text;
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


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
                string FieldValue2 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                int DetailRow2 = 0;
                if (FLAG == "Y")
                {
                    excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
Microsoft.Office.Core.MsoTriState.msoTrue, 340, 680, 200, 80);
                }
                for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    //progressBar1.Value = iRecord;
                    //progressBar1.PerformStep();


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
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
                            DetailRow1 = 26;
                            DetailRow2 = 8;
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

              
                //增加另一talbe處理

                System.Data.DataTable dtmark = GetMenu.GetRmamark(shippingCodeTextBox.Text);
                if (dtmark.Rows.Count > 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 1]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName = "mark";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;

                        DetailRow1++;
                    }
                }


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 1]);
                range.Select();
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                           oMissing);
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                      oMissing);
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1+1, 1]);
                range.Select();
                range.Value2 = "RETURN CARGO WITH NO COMMERCIAL VALUE FOR CUSTOM CLEARANCE ONLY.";
                //增加另二talbe處理

                System.Data.DataTable dtmark1 = GetAuono(shippingCodeTextBox.Text, cardNameTextBox.Text);
     
                for (int a2Row = 0; a2Row <= dtmark1.Rows.Count - 1; a2Row++)
                {

                    //最後一筆不作

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow2, 8]);
                    // range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    string FieldName = "aa";

                    FieldValue2 = "";
                    FieldValue2 = Convert.ToString(dtmark1.Rows[a2Row][FieldName]);

                    range.Value2 = FieldValue2;

                    DetailRow2++;
                }

            }
            finally
            {

             
                     NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);
                
          
                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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
                string Msg = string.Empty;


                System.Diagnostics.Process.Start(NewFileName);
             
                //回應一個下載檔案FileDownload
                // FileUtils.FileDownload(Page, NewFileName);
                this.Cursor = Cursors.Default;//還原預設

            }
        }
        private void GetLADING(string ExcelFile)
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
            excelSheet.Name = shippingCodeTextBox.Text;




            //取得 Excel 的使用區域
            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



            Microsoft.Office.Interop.Excel.Range range = null;

            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;
                string FieldValue1 = string.Empty;
                string FieldValue2 = string.Empty;
                string FieldValue3 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;

                int DetailRow3 = 0;
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

                            DetailRow3 = 24;
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

                //增加另一talbe處理


                System.Data.DataTable mark = Getmark2(shippingCodeTextBox.Text);
                if (mark.Rows.Count > 0)
                {
                    for (int a3Row = 0; a3Row <= mark.Rows.Count - 1; a3Row++)
                    {

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow3, 1]);
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName2 = "ContainerSeals";

                        FieldValue3 = "";
                        FieldValue3 = Convert.ToString(mark.Rows[a3Row][FieldName2]);

                        range.Value2 = FieldValue3;
                        DetailRow3++;
                    }
                }


            }

            finally
            {


               
                    NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                     DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);
          
                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                string Msg = string.Empty;
              
                    System.Diagnostics.Process.Start(NewFileName);
            
            }
        }
        private void  GetExcelProduct4(string ExcelFile,string P1,string FLAG)
        {
            string flag = "Y";
            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;
            excelApp.DisplayAlerts  = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //取得  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
            excelSheet.Name = shippingCodeTextBox.Text;
            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);


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
                string FieldValue2 = string.Empty;
                bool IsDetail = false;
                int DetailRow = 0;
                int DetailRow1 = 0;
                int DetailRow2 = 0;
                if (FLAG == "Y")
                {
                    excelSheet.Shapes.AddPicture(P1, Microsoft.Office.Core.MsoTriState.msoFalse,
    Microsoft.Office.Core.MsoTriState.msoTrue, 340, 680, 200, 80);
                }
                    for (int iRecord = 1; iRecord <= iRowCnt; iRecord++)
                {

                    //progressBar1.Value = iRecord;
                    //progressBar1.PerformStep();


                    for (int iField = 1; iField <= iColCnt; iField++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, iField]);
                        // range.Select();
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
                            DetailRow1 = 49;
                            DetailRow2 = 8;
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

   
                System.Data.DataTable dtmark = GetMenu.GetRmamark(shippingCodeTextBox.Text);
                if (dtmark.Rows.Count > 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                    {

                        //最後一筆不作

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 1]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName = "mark";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;

                        DetailRow1++;
                    }
                }
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 1]);
                range.Select();
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                              oMissing);
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                           oMissing);
                range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                      oMissing);
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1 + 1, 1]);
                range.Select();
                range.Value2 = "RETURN CARGO WITH NO COMMERCIAL VALUE FOR CUSTOM CLEARANCE ONLY.";

                //增加另二talbe處理

                System.Data.DataTable dtmark1 = GetAuono(shippingCodeTextBox.Text, cardNameTextBox.Text);

                for (int a2Row = 0; a2Row <= dtmark1.Rows.Count - 1; a2Row++)
                {

                    //最後一筆不作

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow2, 8]);
                    // range.Select();
                    sTemp = (string)range.Text;
                    sTemp = sTemp.Trim();
                    string FieldName = "aa";

                    FieldValue2 = "";
                    FieldValue2 = Convert.ToString(dtmark1.Rows[a2Row][FieldName]);

                    range.Value2 = FieldValue2;

                    DetailRow2++;
                }

            }
            finally
            {
             
            
                    NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                string Msg = string.Empty;

                System.Diagnostics.Process.Start(NewFileName);


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
        private System.Data.DataTable GetOrderData2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT  b.shippingcode RMANO,c.[add9] as billTo,c.[add10] SHIPTO,'Ship via : '+c.[receiveDay] SHIPBY,c.[shipment] SHIPFROM,Convert(varchar(10),Getdate(),111) as DATE");
            sb.Append("   ,'TO: '+ c.[unloadCargo] SHIPTO2,c.dollarsKind DOLLAR,'SAY TOTAL : US DOLLARS '+c.[add8] TOTAL,");
            sb.Append(" CASE WHEN b.[DIFF]='Y' THEN  cast(seqno+1 as varchar)+')'+b.[INDescription]  ELSE");
            sb.Append("  cast(seqno+1 as varchar)+')'+b.[INDescription]+ char(10) + char(13) +b.[MarkNos]+'  V.'+b.[invoiceno_seq]  END DES");
            sb.Append("  ,b.[InQty] QTY ,b.[UnitPrice] PRICE,b.[Amount] AMOUNT,c.TradeCondition TRADE ");
            sb.Append("                                      FROM [RMA_InvoiceD] as b  ");
            sb.Append("                                      left join RMA_main as c on (b.shippingcode=c.shippingcode)");
            sb.Append(" where b.shippingcode=@shippingcode ");
            sb.Append(" ORDER BY CAST(seqno AS INT) ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));





            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "RMA_invoicem");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        public System.Data.DataTable Getmark2(string shippingcode)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select mark as containerseals from RMA_mark");
            sb.Append(" where shippingcode=@shippingcode");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "mark");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["mark"];
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string P = "Y";
            string USER = fmLogin.LoginID.ToString().ToUpper() + ".JPG";
            string B2 = "//acmew08r2ap//table//SIGN//USER//";
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\RMA\\PACK.xls";

            OrderData = GetOrderData3();
            GetExcelProduct(FileName, B2 + USER, P);
        }
        private System.Data.DataTable GetOrderData5()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select top 1 packageno,seqno,cno from Rma_PackingListd");
            sb.Append(" where shippingcode=@shippingcode and PLNo=@PLNo order by seqno desc ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PLNo", pLNoTextBox.Text));

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
        private System.Data.DataTable GetOrderData3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                                                    SELECT c.DollarsKind DOLAR,'Ref No:'+c.add2 REF,b.SHIPPINGCODE RMANO,c.add7+' PLTS' as PLT,Convert(varchar(10),Getdate(),111) as DATE,c.[receiveDay] SHIPBY,c.[add10] as SHIPTO,c.[shipment] SHIPFROM,");
            sb.Append("                                                         c.[add5] as TNET,c.[add3] TGRO,c.[unloadCargo] SHIPTO2,c.[add9] BILLTO");
            sb.Append("                                                           ,c.[add4] TQTY,c.[add6] TOTAL,b.[PackageNo] PALNO,b.[CNo],");
            sb.Append(" CASE WHEN B.DIFF='Y' THEN b.[DESCGOODS] ");
            sb.Append("  ELSE b.[DESCGOODS]+ char(10) + char(13)+b.[model]+'  V.'+isnull(b.ver,'') END DES");
            sb.Append("                                                           ,b.[Quantity]  QTY ,b.[Net] as NET ,cast(b.[Gross] as varchar) as GRO ,b.[MeasurmentCM] CM ");
            sb.Append("                                                         from   [RMA_PackingListD] as b ");
            sb.Append("                                                           left join RMA_main as c on (b.shippingcode=c.shippingcode)");
            sb.Append("   where b.shippingcode=@shippingcode  ");
            sb.Append("  ORDER BY CAST(seqno AS INT)  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));


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
        private void GetExcelProduct(string ExcelFile,string P1, string FLAG)
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
Microsoft.Office.Core.MsoTriState.msoTrue, 380, 660, 200, 80);
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

                //增加另一talbe處理

                System.Data.DataTable dtmark = GetMenu.GetRmamark(shippingCodeTextBox.Text);
                if (dtmark.Rows.Count > 0)
                {
                    for (int a1Row = 0; a1Row <= dtmark.Rows.Count - 1; a1Row++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[DetailRow1, 6]);
                        // range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        string FieldName = "mark";

                        FieldValue1 = "";
                        FieldValue1 = Convert.ToString(dtmark.Rows[a1Row][FieldName]);

                        range.Value2 = FieldValue1;

                        DetailRow1++;
                    }
                }
        

            }
            finally
            {

            
                    NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);


                try
                {
                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
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

                string Msg = string.Empty;

                System.Diagnostics.Process.Start(NewFileName);
           



            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string aa = cardNameTextBox.Text.Substring(0, 4);
                object[] LookupValues = GetMenu.RmaCardcode(aa);

                if (LookupValues != null)
                {

                    add9TextBox.Text =Convert.ToString(LookupValues[5]) ;
                    add10TextBox.Text = Convert.ToString(LookupValues[6]) ;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void rma_PackingListDDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;
            iRecs = rma_PackingListDDataGridView.Rows.Count - 1;
            e.Row.Cells["dataGridViewTextBoxColumn46"].Value = iRecs.ToString();

        }

        private void rma_MarkDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;
            iRecs = rma_MarkDataGridView.Rows.Count - 1;
            e.Row.Cells["dataGridViewTextBoxColumn14"].Value = iRecs.ToString();

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                
                    OpenFileDialog opdf = new OpenFileDialog();
                    DialogResult result = opdf.ShowDialog();
                    GetExcelContent1(opdf.FileName);
              
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void GetExcelContent1(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

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

            string strText = ""; ;

            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);

            string mailname;
            string address;
            string id;


            for (int i = 1; i <= iRowCnt; i++)
            {

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                id = range.Text.ToString().Trim();
         
                strText = "";
                if (id != "")
                {
                try
                {
                    AddTRACKER_LOG(shippingCodeTextBox.Text, i.ToString(), id);

                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
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
            //可以將 Excel.exe 清除
            System.GC.WaitForPendingFinalizers();
           // rma_DownloadTableAdapter.Fill(rm.Rma_Download, MyID);
            rma_MarkTableAdapter.Fill(rm.Rma_Mark, MyID);
            MessageBox.Show("匯入成功");
        }
        private void AddTRACKER_LOG(string shippingcode, string seq, string mark)
        {



            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Insert into rma_mark(shippingcode,seq,mark) values(@shippingcode,@seq,@mark)");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            command.Parameters.Add(new SqlParameter("@seq", seq));
            command.Parameters.Add(new SqlParameter("@mark", mark));

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

      

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_Mark;
            DataRow newCustomersRow = dt2.NewRow();

            int i = rma_MarkDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["Seq"] = 100;
            newCustomersRow["mark"] = drw["mark"];
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, rma_MarkDataGridView.CurrentRow.Index);
                rma_MarkBindingSource.DataSource = dt2;
                for (int j = 0; j <= rma_MarkDataGridView.Rows.Count - 1; j++)
                {

                    rma_MarkDataGridView.Rows[0].Cells[0].Value = j.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }




        private void rma_InvoiceDDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = rma_InvoiceDDataGridView.Rows.Count - 1;
            e.Row.Cells["dataGridViewTextBoxColumn2"].Value = iRecs.ToString();
            e.Row.Cells["dataGridViewTextBoxColumn43"].Value = "0";
            e.Row.Cells["dataGridViewTextBoxColumn44"].Value = "0";
        }

        private void rma_InvoiceDDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (rma_InvoiceDDataGridView.Columns[e.ColumnIndex].Name == "dataGridViewTextBoxColumn43" ||
                 rma_InvoiceDDataGridView.Columns[e.ColumnIndex].Name == "dataGridViewTextBoxColumn44")
                {
                    decimal iQuantity = 0;
                    decimal iUnitPrice = 0;

                    iQuantity = Convert.ToInt32(this.rma_InvoiceDDataGridView.Rows[e.RowIndex].Cells["dataGridViewTextBoxColumn43"].Value);
                    iUnitPrice = Convert.ToDecimal(this.rma_InvoiceDDataGridView.Rows[e.RowIndex].Cells["dataGridViewTextBoxColumn44"].Value);
                    this.rma_InvoiceDDataGridView.Rows[e.RowIndex].Cells["dataGridViewTextBoxColumn45"].Value = (iQuantity * iUnitPrice).ToString();

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        public static System.Data.DataTable GetAuono(string shippingcode,string aa)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql;
            if (aa == "")
            {
                 sql = "select distinct(rmano) aa from rma_invoiced where shippingcode=@shippingcode";
            }
            else if (aa.Substring(0, 2).ToString() == "友達")
            {
             sql = "select distinct(venderno) aa from rma_invoiced where shippingcode=@shippingcode";
            }
            else
            {
              sql = "select distinct(rmano) aa from rma_invoiced where shippingcode=@shippingcode";
            }
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingcode", shippingcode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "rma_invoiced");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["rma_invoiced"];
        }

    

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {

                if (forecastDayTextBox.Text  != "")
                {
                    GETPACK("01", "", "");
                }
                else
                {

                    System.Data.DataTable dt1 = Getinvoiced(shippingCodeTextBox.Text);
                    System.Data.DataTable dt2 = rm.Rma_PackingListD;

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
                        drw2["model"] = drw["MarkNos"].ToString();
                        drw2["Net"] = NET;
                        drw2["Ver"] = drw["InvoiceNo_seq"];
                        drw2["Quantity"] = drw["InQty"];
                        drw2["DIFF"] = drw["DIFF"];
                        dt2.Rows.Add(drw2);
                    }
                }
          
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        private void GETPACK(string SEQ, string CAR, string CHOSHIP)
        {
  
     
            WriteExcelPACK(SEQ, "", "", CHOSHIP);
            string sbS = "'" + forecastDayTextBox.Text + "'";
            System.Data.DataTable dt3 = util.GetSHIPPACK();

            System.Data.DataTable dt4 = rm.Rma_PackingListD;
            string DPLATENO = "";
            if (dt3.Rows.Count > 0 && rma_PackingListDDataGridView.Rows.Count < 2)
            {
              
                string DESED = "";
                int GV = 0;
                string SERS = "";
                for (int j = 0; j <= dt3.Rows.Count - 1; j++)
                {
                    DataRow drw3 = dt3.Rows[j];
                    DataRow drw2 = dt4.NewRow();
                    string QQ = drw3["QTY"].ToString();
                    string SER = drw3["SER"].ToString();
                    string ES = drw3["ES"].ToString();
                    if (SERS != SER)
                    {
                        GV = 0;
                    }
                    SERS = drw3["SER"].ToString();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["seqno"] = j;
      
                    string PLATENO = drw3["PLATENO"].ToString().Trim();
        
                    string ITEMCODE = drw3["ITEMCODE"].ToString().Trim();
                    string ITEMNAME = drw3["ITEMNAME"].ToString().Trim();
            

                    //if (ITEMCODE == "ACMERMA01.RMA01")
                    //{
                    //    MessageBox.Show("A");
                    //}

           

                    if (String.IsNullOrEmpty(drw2["DescGoods"].ToString()))
                    {


                        System.Data.DataTable H1 = util.GetSHIPPACK3(ITEMCODE);
                        if (H1.Rows.Count > 0)
                        {
                            string MODE = H1.Rows[0][0].ToString().Trim();
                            string GRADE = H1.Rows[0][1].ToString().Trim();
                            string MODE2 = H1.Rows[0][2].ToString().Trim();
                            string VER = H1.Rows[0][3].ToString().Trim();
                            if (MODE.Length > 13)
                            {
                                MODE = MODE.Substring(1, 13);
                            }

                            drw2["DescGoods"] = H1.Rows[0][0].ToString().Trim();
                            drw2["model"] = MODE2;
                            drw2["Ver"] = VER;
                        }
                    }

          

                    if (SER.Trim() != "0")
                    {

                        GV++;
                        if (GV == 1)
                        {
                            System.Data.DataTable dt31 = util.GetSHIPPACK2(SER);
                            if (dt31.Rows.Count > 0)
                            {
                                string PACKAGE = dt31.Rows[0][0].ToString().Trim();
                                if (PACKAGE == "0-0")
                                {
                                    PACKAGE = "";
                                }
                                drw2["PackageNo"] = PACKAGE;
                                drw2["CNo"] = dt31.Rows[0][1].ToString().Trim();
                                drw2["Quantity"] = "'@" + drw3["CARTONQTY"];
                                drw2["Net"] = "'@" + Convert.ToDecimal(drw3["NW"]).ToString("0.000");
                                drw2["Gross"] = "'@" + drw3["GW"];
                                drw2["MeasurmentCM"] = "'@" + drw3["L"] + "x" + drw3["W"] + "x" + drw3["H"];
                            }
                        }
                        if (GV == 2)
                        {
                            System.Data.DataTable dt31 = util.GetSHIPPACK5(SER);
                            if (dt31.Rows.Count > 0)
                            {
                                drw2["Quantity"] = dt31.Rows[0][0].ToString().Trim();
                                drw2["Gross"] = dt31.Rows[0][1].ToString().Trim();
                                drw2["Net"] = dt31.Rows[0][2].ToString().Trim();
                            }
                            drw2["DescGoods"] = "";
                        }
                    }
                    else
                    {
                        GV = 0;

                        if (drw3["ITEMCODE"].ToString().Trim() == "空箱")
                        {

                            drw2["DescGoods"] = "(THIS PALLET INCLUDED " + drw3["CARTONNO"].ToString().Trim() + " EMPTY CARTONS.)";
                            drw2["PackageNo"] = "";
                            drw2["CNo"] = "";
                            drw2["Quantity"] = "";
                            drw2["Net"] = "";
                            drw2["Gross"] = "";
                            drw2["MeasurmentCM"] = "";
                        }
                        else
                        {
                            string PACK = drw3["PLATENO"].ToString().Trim();
                            string CNo = drw3["CARTONNO2"].ToString().Trim();
                            drw2["PackageNo"] = drw3["PLATENO"].ToString().Trim();
                            drw2["CNo"] = drw3["CARTONNO2"].ToString().Trim();
                            drw2["Quantity"] = drw3["CARTONQTY"].ToString().Trim();
                            drw2["Net"] = Convert.ToDecimal(drw3["NW"]).ToString("#,##0.000");
                            drw2["Gross"] = Convert.ToDecimal(drw3["GW"]).ToString("#,##0.00");
                            if (!String.IsNullOrEmpty(drw3["L"].ToString()))
                            {
                                drw2["MeasurmentCM"] = drw3["L"] + "x" + drw3["W"] + "x" + drw3["H"];
                            }
                        }
                    }

                    string DESE = drw2["DescGoods"].ToString();
                    //int n;
                    //if (int.TryParse(drw2["Quantity"].ToString(), out n) && int.TryParse(drw3["QTY"].ToString(), out n))
                    //{
                    //    if (DESE != DESED && ACME == -1)
                    //    {
                    //        int QTY = Convert.ToInt16(drw2["Quantity"]);
                    //        int QTY2 = Convert.ToInt16(drw3["QTY"]);
                    //        if (QTY >= QTY2)
                    //        {
                    //            drw2["PALQTY"] = drw3["QTY"].ToString();
                    //        }

                    //        //20180604
                    //        System.Data.DataTable G11 = util.GetSHIPPACKQTY(ITEMCODE);
                    //        if (G11.Rows.Count > 0)
                    //        {
                    //            drw2["PALQTY"] = G11.Rows[0][0].ToString();
                    //        }
                    //    }
                    //}
                    //if (GV == 1)
                    //{
                    //    if (DESE != DESED)
                    //    {
                    //        drw2["PALQTY"] = drw3["QTY"].ToString();
                    //    }
                    //}
                    //if (GV == 2)
                    //{
                    //    drw2["PALQTY"] = "";
                    //}
                    DESED = DESE;

                    //if (!checkBox6.Checked)
                    //{
                    //    if (DPLATENO == PLATENO)
                    //    {
                    //        drw2["PackageNo"] = "";
                    //    }
                    //}

                    if (!String.IsNullOrEmpty(PLATENO))
                    {
                        DPLATENO = PLATENO;
                    }
                    if (GV <= 2)
                    {
                        dt4.Rows.Add(drw2);
                    }


                }

            }



        }
        private void WriteExcelPACK(string SEQ, string CHE, string CAR, string CHOSHIP)
        {

            util.DELPACK();

            int SQ = Convert.ToInt16(SEQ);
       
            StringBuilder sb = new StringBuilder();
            string SHIPPINGCODE = forecastDayTextBox.Text;
            int M1 = 0;
       
            string BLC = "";
         
        
            System.Data.DataTable dt3 = util.GetWHPACK(SHIPPINGCODE, "", "", "", "");
            if (dt3.Rows.Count == 0)
            {
                MessageBox.Show("包裝明細無資料");
                return;
            }
            if (dt3.Rows.Count > 0)
            {
                string PLATENO;
                string PLATENOE1 = "";
                string PLATENOE2 = "";
                string CARTONNO;
                string ITEMCODE;
                string QTY;
                string CARTONQTY;
                string NW;
                string GW;
                string L;
                string W;
                string H;
                string LOACTION;
                string PLATENO2 = "";
                string L2 = "";
                string W2 = "";
                string H2 = "";
                string MITEM = "";
                string GW2 = "";
                string INVOICE = "";
                string ITEMNAME = "";
                string WHNO = "";
                string WHNOD = "";
                int SER = 0;
                int SER2 = 0;
                string SERX = "";
                int CARTONNO2 = 0;
                string ES;
                string CARTONNO3 = "";
                string CARTONNO5 = "";
                int SER3 = 0;

                for (int j = 0; j <= dt3.Rows.Count - 1; j++)
                {
                    DataRow drw3 = dt3.Rows[j];


                    WHNO = drw3["SHIPPINGCODE"].ToString().Trim();
                    PLATENO = drw3["PLATENO"].ToString().Trim();
                    PLATENO2 = drw3["PLATENO2"].ToString().Trim();
                    CARTONNO = drw3["CARTONNO"].ToString().Trim();
                    ITEMCODE = drw3["ITEMCODE"].ToString().Trim();
                    ITEMNAME = drw3["ITEMNAME"].ToString().Trim();
                    QTY = drw3["QTY"].ToString().Trim();
                    CARTONQTY = drw3["CARTONQTY"].ToString().Trim();
                    NW = drw3["NW"].ToString().Trim();
                    GW = drw3["GW2"].ToString().Trim();
                    L = drw3["L"].ToString().Trim();
                    W = drw3["W"].ToString().Trim();
                    H = drw3["H"].ToString().Trim();

                    if (j == 0)
                    {
                        PLATENOE1 = PLATENO;
                    }
                    if (j == 1)
                    {
                        PLATENOE2 = PLATENO;
                    }
                    CARTONNO5 = CARTONNO;
                    if (!String.IsNullOrEmpty(PLATENO2))
                    {
                        System.Data.DataTable H1 = util.GetSHIPPACK9(WHNO, PLATENO2);
                        if (H1.Rows.Count > 0)
                        {
                            CARTONNO5 = H1.Rows[0][0].ToString();
                        }
                    }
                    LOACTION = drw3["LOACTION"].ToString().Trim();
                    INVOICE = drw3["AUNO"].ToString().Trim();
                    ES = drw3["ES"].ToString().Trim();
                    if (QTY == "空箱")
                    {
                        QTY = "0";
                        ITEMCODE = "空箱";
                    }

                    int CARTONNO4 = 0;
                    if (WHNOD != WHNO)
                    {
                        CARTONNO2 = 0;
                    }
                    if (cardCodeTextBox.Text == "1362-00")
                    {
                        if (PLATENOE1 == "1" && PLATENOE2 == "1")
                        {

                        }
                        else
                        {
                            if (PLATENO == "1")
                            {
                                CARTONNO2 = 0;
                            }
                        }
                    }
                    if (!String.IsNullOrEmpty(ITEMCODE))
                    {
                        if (ITEMCODE != "空箱")
                        {
                            CARTONNO4 = CARTONNO2 + 1;
                            if (String.IsNullOrEmpty(CARTONNO))
                            {
                                CARTONNO = "0";
                            }
                            CARTONNO2 += Convert.ToInt16(CARTONNO);
                        }

                        //if (ITEMCODE == "M270DAN02.55QA2")
                        //{
                        //    MessageBox.Show("A");
                        //}

                        string F1 = CARTONNO5 + ITEMCODE + GW + L + W + H + QTY;
                        if ((CARTONNO5 + ITEMCODE + GW + L + W + H + QTY != MITEM) || (String.IsNullOrEmpty(L)))
                        {
                            SERX = SER2.ToString();
                            SER3 = 0;
                        }
                        else
                        {
                            if (SER3 == 0)
                            {
                                SER++;
                                SERX = SER.ToString();
                                util.UPPACKS(SERX);
                                SER3 = 1;
                            }

                        }
                        if (!String.IsNullOrEmpty(PLATENO2))
                        {
                            MITEM = CARTONNO5 + ITEMCODE + GW + L + W + H + QTY;
                        }
                        else
                        {
                            MITEM = CARTONNO + ITEMCODE + GW + L + W + H + QTY;
                        }


                        CARTONNO3 = CARTONNO4.ToString().Trim() + "~" + CARTONNO2.ToString().Trim();
                        if (CARTONNO == "1")
                        {
                            CARTONNO3 = CARTONNO4.ToString();
                        }
                        if (CARTONNO == "0")
                        {
                            CARTONNO3 = "";
                        }
                        if (String.IsNullOrEmpty(NW))
                        {
                            System.Data.DataTable G1 = util.GetSHIPPACK6(ITEMCODE, QTY);
                            if (G1.Rows.Count > 0)
                            {
                                NW = G1.Rows[0][0].ToString();
                            }
                            else
                            {
                                System.Data.DataTable G2 = util.GetSHIPPACK7(ITEMCODE);
                                if (G2.Rows.Count > 0)
                                {
                                    string PAL_NW = G2.Rows[0]["PAL_NW"].ToString();
                                    string PAL_QTY = G2.Rows[0]["PAL_QTY"].ToString();

                                    decimal n;
                                    if (decimal.TryParse(PAL_NW, out n) && decimal.TryParse(PAL_QTY, out n) && decimal.TryParse(QTY, out n))
                                    {
                                        NW = ((Convert.ToDecimal(PAL_NW) / Convert.ToDecimal(PAL_QTY)) * Convert.ToDecimal(QTY)).ToString("#,##0.000");
                                    }
                                }
                            }
                        }
                        else
                        {
                            NW = Convert.ToDecimal(NW).ToString("0.000");
                        }
                        util.AddPACK(PLATENO, CARTONNO, ITEMCODE, QTY, CARTONQTY, NW, GW, L, W, H, LOACTION, SERX, CARTONNO3, INVOICE, ITEMNAME, WHNO, ES);
                    }

                    WHNOD = WHNO;
                }
            }



        }
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_InvoiceD;
            DataRow newCustomersRow = dt2.NewRow();

            int i = rma_InvoiceDDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["Seqno"] = "100";
            newCustomersRow["INDescription"] = drw["INDescription"];
            newCustomersRow["MarkNos"] = drw["MarkNos"];
            newCustomersRow["InvoiceNo_seq"] = drw["InvoiceNo_seq"];
            newCustomersRow["InQty"] = drw["InQty"];
            newCustomersRow["UnitPrice"] = drw["UnitPrice"];
            newCustomersRow["Amount"] = drw["Amount"];
            newCustomersRow["RmaNo"] = drw["RmaNo"];
            newCustomersRow["VenderNo"] = drw["VenderNo"];
            newCustomersRow["CodeName"] = drw["CodeName"];
            newCustomersRow["Grade"] = drw["Grade"];
          
     
            try
            {
                dt2.Rows.InsertAt(newCustomersRow, rma_InvoiceDDataGridView.Rows.Count);
                rma_InvoiceDBindingSource1.DataSource = dt2;


                for (int j = 0; j <= rma_InvoiceDDataGridView.Rows.Count - 2; j++)
                {
                    rma_InvoiceDDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_PackingListD;
            DataRow newCustomersRow = dt2.NewRow();

            int i = rma_PackingListDDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["SeqNo"] = "100";
            newCustomersRow["PackageNo"] = drw["PackageNo"];
            newCustomersRow["CNo"] = drw["CNo"];
            newCustomersRow["DescGoods"] = drw["DescGoods"];
            newCustomersRow["model"] = drw["model"];
            newCustomersRow["Ver"] = drw["Ver"];
            newCustomersRow["Quantity"] = drw["Quantity"];
            newCustomersRow["Net"] = drw["Net"];
            newCustomersRow["Gross"] = drw["Gross"];
            newCustomersRow["MeasurmentCM"] = drw["MeasurmentCM"];

            try
            {
                dt2.Rows.InsertAt(newCustomersRow, rma_PackingListDDataGridView.Rows.Count);
                rma_PackingListDBindingSource.DataSource = dt2;


                for (int j = 0; j <= rma_PackingListDDataGridView.Rows.Count - 2; j++)
                {
                    rma_PackingListDDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }
 
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void 插入列ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_PackingListD;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;


            try
            {

                dt2.Rows.InsertAt(newCustomersRow, rma_PackingListDDataGridView.CurrentRow.Index);
                rma_PackingListDBindingSource.DataSource = dt2;

                for (int j = 0; j <= rma_PackingListDDataGridView.Rows.Count-2; j++)
                {
                    rma_PackingListDDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void 插入列ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_InvoiceD;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["SeqNo"] = 100;
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, rma_InvoiceDDataGridView.CurrentRow.Index);
                rma_InvoiceDBindingSource1.DataSource = dt2;

                for (int j = 0; j <= rma_InvoiceDDataGridView.Rows.Count - 2; j++)
                {
                    rma_InvoiceDDataGridView.Rows[j].Cells[0].Value = j.ToString();
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
            string sql = "select * from rma_invoiced where shippingcode=@shippingcode ";

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
        public static System.Data.DataTable Getmain(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select add3,add7 from Rma_main where shippingcode=@shippingcode  ";

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
        public static System.Data.DataTable GetMark(string shippingcode)
        {
            SqlConnection connection = globals.Connection;
            string sql = "select * from RMA_MARK where shippingcode=@shippingcode order by cast(seq as int) ";

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
        private void button8_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            FileName = lsAppDir + "\\Excel\\RMA\\BL.xls";

            OrderData = GetOrderData();
            GetLADING(FileName);
        }

        private void rMA_LADINGDDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            
                int iRecs;
                iRecs = rMA_LADINGDDataGridView.Rows.Count - 1;
                e.Row.Cells["dataGridViewTextBoxColumn59"].Value = iRecs.ToString();
        }


        private System.Data.DataTable GetOrderData()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select T0.ShippingCode 單號,add9 BILLTO,add10 SHIPTO,receivePlace 收貨地,goalPlace 目的地");
            sb.Append(" ,shipping_OBU OCEAN,T1.Packages MARKS,Description DES,Cargo CARGO,Measurement MEA,bRAND SO,buCntctPrsn FEIGHT from rma_main T0");
            sb.Append(" LEFT JOIN RMA_LADINGD T1 ON (T0.ShippingCode=T1.ShippingCode)");
            sb.Append(" where t0.shippingcode=@shippingcode  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));

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

        private System.Data.DataTable GetNET(string MODEL,string VER)
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


        private System.Data.DataTable GetDESC(string U_TMODEL)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT U_ITEMNAME FROM OITM  WHERE U_TMODEL=@U_TMODEL AND ISNULL(U_ITEMNAME,'') <> ''");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@U_TMODEL", U_TMODEL));
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
        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt1 = GetMark(shippingCodeTextBox.Text);
                System.Data.DataTable dt3 = Getmain(shippingCodeTextBox.Text);
                
                System.Data.DataTable dt2 = rm.RMA_LADINGD;

                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw3 = dt3.Rows[0];
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = drw["Seq"];
                    drw2["ContainerSeals"] = drw["Mark"];
                    if (i.ToString() == "0")
                    {
                        drw2["Packages"] = drw3["add7"];
                        drw2["Description"] = "AUO BRAND";
                        drw2["Cargo"] = drw3["add3"];
                        
                    }
                    if (i.ToString() == "1")
                    {
                        drw2["Description"] = "TFT LCD MODULE";
                    }
                    if (i.ToString() == "2")
                    {
                        drw2["Description"] = "INVOICE NO:"+ shippingCodeTextBox.Text.ToString();
                    }
                    dt2.Rows.Add(drw2);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

  

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt1 = Getinvoiced(shippingCodeTextBox.Text);
                System.Data.DataTable dt2 = rm.Rma_Mark2;

                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    
                    drw2["MODEL"] = drw["MarkNos"];
                    drw2["VER"] = drw["InvoiceNo_seq"];
                    drw2["QTY"] = drw["InQty"];
                    drw2["CARDNAME"] = drw["CodeName"];
                    drw2["RMANO"] = drw["RmaNo"].ToString().Trim();
                    drw2["VENDER"] = drw["VENDER"];


                    dt2.Rows.Add(drw2);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\SZMARK.xls";

                //取得 Excel 資料
                OrderData = SZMARK();
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //產生 Excel Report
                ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable SZMARK()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select RMANO,CARDNAME,''''+VENDER VENDER,MODEL,VER,QTY from Rma_Mark2 where shippingcode=@shippingcode ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@shippingcode", shippingCodeTextBox.Text));
         
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "RMA_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void rma_PackingListDDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if(!checkBox1.Checked)
            {
                if (rma_PackingListDDataGridView.Columns[e.ColumnIndex].Name == "Quantity" ||
             rma_PackingListDDataGridView.Columns[e.ColumnIndex].Name == "Net")
                {
                    string MODEL = rma_PackingListDDataGridView.Rows[e.RowIndex].Cells["model"].Value.ToString();
                    string VER = rma_PackingListDDataGridView.Rows[e.RowIndex].Cells["Ver"].Value.ToString();
                    string Description = rma_PackingListDDataGridView.Rows[e.RowIndex].Cells["Description"].Value.ToString().ToUpper();
                    string Quantity = rma_PackingListDDataGridView.Rows[e.RowIndex].Cells["Quantity"].Value.ToString();
                    string Net = rma_PackingListDDataGridView.Rows[e.RowIndex].Cells["Net"].Value.ToString();

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
                                    this.rma_PackingListDDataGridView.Rows[e.RowIndex].Cells["Net"].Value = (H1 * QTY2).ToString().Replace(".0", "");
                                }
                            }
                        }
                    }

                }
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                System.Data.DataTable dt1 = Getinvoiced(shippingCodeTextBox.Text);
                System.Data.DataTable dt2 = rm.Rma_InvoiceD2;

                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
       
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;
                    drw2["SeqNo"] = drw["SeqNo"];
                    drw2["model"] = drw["MarkNos"].ToString();
                    drw2["VER"] = drw["InvoiceNo_seq"];
                    drw2["QTY"] = drw["InQty"];
                    drw2["RmaNo"] = drw["RmaNo"];
                    drw2["VenderNo"] = drw["VenderNo"];
                    drw2["CUST"] = drw["CodeName"];
                    drw2["Grade"] = drw["Grade"];
                    drw2["VENDER"] = drw["VENDER"];
                    drw2["LOCATION"] = dollarsKindTextBox.Text;
                    
              
                    dt2.Rows.Add(drw2);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            createDateTextBox.Text = comboBox1.Text;
        }

        private void button15_Click(object sender, EventArgs e)
        {

            string F1 = createDateTextBox.Text + "收貨通知單---" + shippingCodeTextBox.Text;


            try
            {
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\收貨工單.xls";


                //Excel的樣版檔
                string ExcelTemplate = FileName;
                //香港倉-宏高收貨通知單---RMA20181019001X---100PCS
                string F2 = createDateTextBox.Text + "收貨通知單---" + shippingCodeTextBox.Text + "---" + GetOrderDataN2().Rows[0][0].ToString()+"PCS.xls";
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +F2;

                //產生 Excel Report
                ExcelReport.ExcelReportOutput(GetOrderDataN(F1), ExcelTemplate, OutPutFile, "N");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable GetOrderDataN(string F1)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ''''+VenderNo VRMANO, ROW_NUMBER() OVER(ORDER BY RMANO) AS  LINE,'PCS' PCS,T0.BoardCountNO,");
            sb.Append(" CASE WHEN BoardCountNo ='進口' THEN  SUBSTRING(T0.arriveDay,1,4)+'/'+SUBSTRING(T0.arriveDay,5,2)+'/'+SUBSTRING(T0.arriveDay,7,2) ");
            sb.Append(" WHEN BoardCountNo ='出口' THEN  SUBSTRING(T0.Quantity,1,4)+'/'+SUBSTRING(T0.Quantity,5,2)+'/'+SUBSTRING(T0.Quantity,7,2) COLLATE  Chinese_Taiwan_Stroke_CI_AS END DNOW2");
            sb.Append(" ,convert(varchar, getdate(),111) DNOW,VENDER,RMANO,CUST,MODEL,VER,GRADE,QTY,'' INVOICE, LOCATION,@TITLE TITLE,@LGOIN LGOIN");
            sb.Append(" FROM RMA_MAIN T0 INNER JOIN Rma_InvoiceD2  T1 ON (T0.ShippingCode =T1.ShippingCode) WHERE T0.ShippingCode=@ShippingCode ORDER BY   RmaNo   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TITLE", F1));
            command.Parameters.Add(new SqlParameter("@LGOIN", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
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
        private System.Data.DataTable GetOrderDataN2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(SUM(QTY),0) QTY FROM Rma_InvoiceD2 WHERE    ShippingCode=@ShippingCode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
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
        private void button16_Click(object sender, EventArgs e)
        {

            RmaNo frm1 = new RmaNo();
            frm1.q1 = "進金生";

            if (frm1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    System.Data.DataTable dt1 = GetAR2("進金生",frm1.q);
                    System.Data.DataTable dt2 = rm.Rma_InvoiceD2;

                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["RmaNo"] = drw["U_RMA_NO"];
                        drw2["model"] = drw["U_RMODEL"];
                        drw2["SeqNo"] = "0";
                        drw2["VenderNo"] = drw["U_AUO_RMA_NO"];

                        //string T1 = drw["aa"].ToString();
                        //string T2 = drw["U_RMODEL"].ToString();

                        //string SIZE2 = "";
                        //int I2 = T2.ToUpper().IndexOf("_OPEN CELL");
                        //int I3 = T2.ToUpper().IndexOf("KIT_");
                        //int I4 = T2.ToUpper().IndexOf("KIT AD BOARD");
                        //int I5 = T2.ToUpper().IndexOf("KIT AD INVERTER");
                        //int I6 = T2.ToUpper().IndexOf("KIT DRIVER BOARD");
                        //int I7 = T2.ToUpper().IndexOf("KIT INVERTER");
                        //int I8 = T2.ToUpper().IndexOf("_T CON");
                        //int I12 = T2.ToLower().IndexOf("enclosure");

                        ////OpenFrame 
                        //int I9 = T2.ToUpper().IndexOf("\"");
                        //int I11 = T2.ToUpper().IndexOf("”");
                        //int I10 = T2.ToUpper().IndexOf("OPENFRAME");
                        //if (I9 != -1)
                        //{
                        //    SIZE2 = T2.Substring(0, I9);
                        //}
                        //if (I11 != -1)
                        //{
                        //    SIZE2 = T2.Substring(0, I11);
                        //}


                        //if (I2 != -1)
                        //{

                        //    drw2["model"] = drw["U_RMODEL"].ToString().Substring(0, I2);


                        //}
                        //else if (I8 != -1)
                        //{
                        //    drw2["model"] = drw["U_RMODEL"].ToString().Substring(0, I8);

                        //}
                        //else if (I3 != -1)
                        //{
                        //    drw2["model"] = SIZE2 + "\" KIT";
                        //}
                        //else if (I4 != -1)
                        //{
                        //    drw2["model"] = SIZE2 + "\" KIT_AD Board";
                        //}
                        //else if (I5 != -1)
                        //{
                        //    drw2["model"] = SIZE2 + "\" KIT_Inverter";
                        //}
                        //else if (I6 != -1)
                        //{
                        //    drw2["model"] = SIZE2 + "\" KIT_Driver Board";
                        //}
                        //else if (I7 != -1)
                        //{
                        //    drw2["model"] = SIZE2 + "\" KIT_Inverter";
                        //}
                        //else if (I10 != -1)
                        //{
                        //    drw2["model"] = SIZE2 + "\" OpenFrame";
                        //}
                        //else if (I12 != -1)
                        //{
                        //    drw2["model"] = SIZE2 + "\" Enclosure Display";
                        //}
                        //else
                        //{

                        //    string MODEL = drw["U_RMODEL"].ToString();
                        //    drw2["model"] = MODEL;
           
                        //}

                  
                        int g = drw["U_Rquinity"].ToString().IndexOf("+");
                        int t = drw["U_Rquinity"].ToString().LastIndexOf("+");
                        string h;
                        string s;


                        if (g == -1)
                        {
                            s = drw["U_Rquinity"].ToString();
                            drw2["QTY"] = s;
                        }
                        else
                        {
                            s = drw["U_Rquinity"].ToString().Substring(g + 1);

                            try
                            {

                                if (drw["U_Rquinity"].ToString().Substring(3, 1) != "+")
                                {
                                    h = drw["U_Rquinity"].ToString().Substring(0, 2);
                                }
                                else
                                {
                                    h = drw["U_Rquinity"].ToString().Substring(0, 1);
                                }
                                int a = Convert.ToInt16(s.ToString());
                                int b = Convert.ToInt16(h.ToString());
                                drw2["QTY"] = (a + b).ToString();
                            }
                            catch (Exception ex)
                            {
                                h = drw["U_Rquinity"].ToString().Substring(0, 1);
                                drw2["QTY"] = h.ToString();
                            }

                        }

                        drw2["Grade"] = drw["U_Rgrade"];
                        drw2["VER"] = drw["U_Rver"];

                        drw2["CUST"] = drw["U_cusname_s"];
                        drw2["VENDER"] = drw["U_Rvender"];
                        drw2["LOCATION"] = dollarsKindTextBox.Text;
                        dt2.Rows.Add(drw2);
                    }

                    for (int j = 0; j <= rma_InvoiceDDataGridView.Rows.Count - 2; j++)
                    {
                        rma_InvoiceDDataGridView.Rows[j].Cells[0].Value = j.ToString();
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

        private void rma_InvoiceDDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void btnEmailPeter_Click(object sender, EventArgs e)
        {
            string template;
            StreamReader objReader;
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            

            FileName = lsAppDir + "\\MailTemplates\\RmaToPeter.html";
            objReader = new StreamReader(FileName);

            template = objReader.ReadToEnd();
            objReader.Close();
            objReader.Dispose();

            string Customer = "";
            string htmlRmaNo = "";//要放信件##RmaNo##的

            string RmaNo = "";

            StringWriter writer = new StringWriter();
            HtmlTextWriter htmlWriter = new HtmlTextWriter(writer);
            string TOTAL = "";
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<table class='GridBorder'  border='1' cellspacing='0' rules='all'  style='border:3px #F5F5DC groove;'>");

            //欄位
            sb.AppendLine("<tr>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">RMA NO</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">Customer</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">Model</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">Version</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">Qty</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">MADE IN</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">Carton no.</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">材積</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">重量</font></th>");
            sb.AppendLine("<th bgcolor =\"#272727\"><font color=\"white\">Remark</font></th>");
            sb.AppendLine("</tr>");

            foreach (DataGridViewRow dgvrow in rma_InvoiceDDataGridView.Rows) 
            {
                if (dgvrow.Cells[0].Value != "" && dgvrow.Cells[0].Value != null) 
                {
                    RmaNo += dgvrow.Cells["dataGridViewTextBoxColumn47"].Value.ToString() + ";";
                }
            }
            RmaNo = RmaNo.TrimEnd(';');
            System.Data.DataTable dt = GetRmaData(RmaNo, dollarsKindTextBox.Text.Replace("MADE IN",""));
            int i = 0;
            foreach (DataRow rows in dt.Rows)
            {
                //不同行分色
                if (i % 2 == 0)
                {
                    sb.AppendLine("<tr bgcolor=\"#C0C0C0\">");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        sb.AppendLine("<td>" + rows[j] + "</td>");
                        if (j == 1 && rows[j] != "")
                        {
                            Customer = rows[j].ToString();
                        }
                        else if (j == 0) 
                        {
                            htmlRmaNo += rows[j].ToString() +"、";
                        }
                    }
                    sb.AppendLine("</tr>");
                }
                else
                {
                    sb.AppendLine("<tr bgcolor=\"#E5E4E2\">");
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        sb.AppendLine("<td>" + rows[j] + "</td>");
                        if (j == 1 && rows[j] != "")
                        {
                            Customer = rows[j].ToString();
                        }
                        else if (j == 0)
                        {
                            htmlRmaNo += rows[j].ToString() + "、";
                        }
                    }
                    sb.AppendLine("</tr>");
                }
                i++;
            }
            TOTAL = dt.Rows[dt.Rows.Count - 1]["InQty"].ToString();

            sb.AppendLine("</table>");
            htmlRmaNo = htmlRmaNo.TrimEnd('、');

            template = template.Replace("##Customer##", Customer);

            template = template.Replace("##RmaNo##", htmlRmaNo);

            template = template.Replace("##Template##", sb.ToString());

            string SlpName = globals.UserID;

            string MailToAddress = "";

            string strSubject = "請協助打包_" + Customer + "_" + htmlRmaNo + "_TTL: "+ TOTAL + " PCS ";


            MailToAddress = "peterdu@acmepoint.com" + ";" + "federliu@acmepoint.com";
            //MailToAddress = "nesschou@acmepoint.com;";


            MailToAddress = MailToAddress.TrimEnd(';');

            string MailFromAddress = "workflow@acmepoint.com";

            MailToPD(strSubject, MailFromAddress, MailToAddress, template);


        }
        private void MailToPD(string strSubject, string MailFromAddress, string MailToAddress, string MailContent)
        {
            MailMessage message = new MailMessage();

            message.From = new MailAddress(MailFromAddress, "系統發送");
            string[] MailToAdd = MailToAddress.Split(';');
            foreach (string add in MailToAdd)
            {
                message.To.Add(new MailAddress(add));
            }




            string myMailEncoding = "utf-8";
            message.Subject = strSubject;
            //message.Body = string.Format("<html><body><P>Dear {0},</P><P>請參考!</P> {1} </body></html>", SlpName, MailContent);
            message.Body = MailContent;
            //格式為 Html
            message.IsBodyHtml = true;

            SmtpClient client = new SmtpClient();
            client.Host = "ms.mailcloud.com.tw";
            client.UseDefaultCredentials = true;

            //string pwd = "Y4/45Jh6O4ldH1CvcyXKig==";
            //pwd = Decrypt(pwd, "1234");

            string pwd = "@cmeworkflow";

            //client.Credentials = new System.Net.NetworkCredential("TerryLee@acmepoint.com", pwd);
            client.Credentials = new System.Net.NetworkCredential("workflow@acmepoint.com", pwd);
            try
            {
                client.Send(message);
                MessageBox.Show("信件已寄出");
            }
            catch (SmtpFailedRecipientsException ex)
            {
                for (int i = 0; i < ex.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = ex.InnerExceptions[i].StatusCode;
                    if (status == SmtpStatusCode.MailboxBusy ||
                        status == SmtpStatusCode.MailboxUnavailable)
                    {
                        //SetMsg("Delivery failed - retrying in 5 seconds.");
                        System.Threading.Thread.Sleep(5000);
                        client.Send(message);

                    }
                    else
                    {
                        //SetMsg(String.Format("Failed to deliver message to {0}",
                        // ex.InnerExceptions[i].FailedRecipient));
                    }
                }
            }
            catch (Exception ex)
            {
                //SetMsg(String.Format("Exception caught in RetryIfBusy(): {0}",
                // ex.ToString()));
            }

        }
        private System.Data.DataTable GetRmaData(string RmaNo,string dollarsKind)
        {

            SqlConnection connection = globals.Connection;

            string[] RMANO = RmaNo.Split(';');
            StringBuilder sb = new StringBuilder();
            sb.Append(" select RmaNo,CodeName,MarkNos,InvoiceNo_seq,InQty,'CHINA','','','',''");
            sb.Append(" from [Rma_InvoiceD]");
            sb.Append(" WHERE (");

            for (int i = 0; i < RMANO.Length; i++) 
            {
                if (i == 0)
                {
                    sb.Append(" RmaNo like '" + RMANO[i] + "'");
                }
                else 
                {
                    sb.Append("or RmaNo like '" + RMANO[i] + "'");
                }

            }
            sb.Append(" )");

            sb.Append(" UNION");

            sb.Append(" select '','','','TTL',SUM(cast(InQty as int)),'','','','',''");
            sb.Append(" from Rma_InvoiceD");
            sb.Append(" WHERE (");

            for (int i = 0; i < RMANO.Length; i++)
            {
                if (i == 0)
                {
                    sb.Append(" RmaNo like '" + RMANO[i] + "'");
                }
                else
                {
                    sb.Append("or RmaNo like '" + RMANO[i] + "'");
                }

            }
            sb.Append(" )");

            sb.Append(" order by RmaNo desc");
           
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.Parameters.Add(new SqlParameter("@dollarsKind", dollarsKind));
            command.CommandType = CommandType.Text;
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
    }
}


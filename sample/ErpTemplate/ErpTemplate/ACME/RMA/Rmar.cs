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
namespace ACME
{
    public partial class Rmar : ACME.fmBase1
    {
        string H1 = "";
        string NAME = "";
        string PHONE = "";
        private System.Data.DataTable OrderData1;
        private System.Data.DataTable OrderData2;
        string NewFileName = "";
        string PCS = "";
        string CART = "";
        Attachment data = null;
        string GlobalMailContent = "";
        public Rmar()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            rma_mainrTableAdapter.Connection = MyConnection;
            rma_InvoiceDrTableAdapter.Connection = MyConnection;
            rma_DELIVERYTableAdapter.Connection = MyConnection;

        }
        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
        }



        public override void EndEdit()
        {
            Control();
        }
        public override void AfterEndEdit()
        {
            System.Data.DataTable dt1 = rm.Rma_InvoiceDr;
            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DateTime R1 = Convert.ToDateTime(GetMenu.DayS(forecastDayTextBox.Text));
                DataRow drw = dt1.Rows[i];
                string aa = drw["RmaNo"].ToString();
                string bb = drw["shippingcode"].ToString();
                string Qty3 = drw["Qty3"].ToString();
                string Qty4 = drw["Qty4"].ToString();

                if (string.IsNullOrEmpty(Qty3))
                {
                    UPDATEJOBNO(bb, R1, Qty4, aa);
                }
            }
        }
        public override void AfterEdit()
        {


            shippingCodeTextBox.ReadOnly = true;


        }
        public override void AfterCancelEdit()
        {
            Control();
        }
        public override void AfterAddNew()
        {
            Control();
        }
        private void Control()
        {
            shippingCodeTextBox.ReadOnly = true;
            button1.Enabled = true;
            button10.Enabled = true;
            button14.Enabled = true;
            button15.Enabled = true;
            button16.Enabled = true;
            button13.Enabled = true;
            button23.Enabled = true;
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            button19.Enabled = true;
            button21.Enabled = true;
            button22.Enabled = true;


        }
        public override void SetInit()
        {

            MyBS = rma_mainrBindingSource;
            MyTableName = "Rma_mainr";
            MyIDFieldName = "ShippingCode";
            UtilSimple.SetLookupBinding(add6ComboBox, "add6", rma_mainrBindingSource, "add6");
            UtilSimple.SetLookupBinding(tradeConditionComboBox, "TradeCondition", rma_mainrBindingSource, "TradeCondition");
            UtilSimple.SetLookupBinding(closeDayComboBox, "closeDay", rma_mainrBindingSource, "closeDay");
            UtilSimple.SetLookupBinding(comboBox1, GetMenu.GetBU("RMAR"), "DataText", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, GetMenu.GetBU("RMAR"), "DataText", "DataValue");


            //³B²z½Æ»s
            MasterTable = rm.Rma_mainr;
            DetailTables = new System.Data.DataTable[] { rm.Rma_InvoiceDr };
            DetailBindingSources = new BindingSource[] { rma_InvoiceDrBindingSource };

        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "RMR" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;
            kyes = this.shippingCodeTextBox.Text;
            this.rma_mainrBindingSource.EndEdit();
            kyes = null;
        }
        public override void AfterCopy()
        {
            if (kyes == null)
            {
                string NumberName = "RMR" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                this.shippingCodeTextBox.Text = NumberName + AutoNum + "X";
                kyes = this.shippingCodeTextBox.Text;
            }
        }
        public override void FillData()
        {
            try
            {
                rma_mainrTableAdapter.Fill(rm.Rma_mainr, MyID);
                rma_InvoiceDrTableAdapter.Fill(rm.Rma_InvoiceDr, MyID);
                rma_DELIVERYTableAdapter.Fill(rm.Rma_DELIVERY, MyID);

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

                rma_InvoiceDrBindingSource.MoveFirst();

                for (int i = 0; i <= rma_InvoiceDrBindingSource.Count - 1; i++)
                {
                    DataRowView row3 = (DataRowView)rma_InvoiceDrBindingSource.Current;

                    row3["SeqNo"] = i;



                    rma_InvoiceDrBindingSource.EndEdit();

                    rma_InvoiceDrBindingSource.MoveNext();
                }


                rma_DELIVERYBindingSource.MoveFirst();

                for (int i = 0; i <= rma_DELIVERYBindingSource.Count - 1; i++)
                {
                    DataRowView row4 = (DataRowView)rma_DELIVERYBindingSource.Current;

                    row4["SeqNo"] = i;

                    rma_DELIVERYBindingSource.EndEdit();

                    rma_DELIVERYBindingSource.MoveNext();
                }

                rma_mainrTableAdapter.Connection.Open();

                rma_mainrBindingSource.EndEdit();
                rma_InvoiceDrBindingSource.EndEdit();
                rma_InvoiceDrBindingSource.EndEdit();

                rma_mainrTableAdapter.Update(rm.Rma_mainr);
                rma_InvoiceDrTableAdapter.Update(rm.Rma_InvoiceDr);
                rma_DELIVERYTableAdapter.Update(rm.Rma_DELIVERY);

                UpdateData = true;
            }
            catch (Exception ex)
            {


                MessageBox.Show(ex.Message, "§ó·s¿ù»~", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.rma_mainrTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
        public void UPDATEJOBNO(string u_jobno, DateTime U_ACME_BackDate1, string u_acme_backqty1, string u_rma_no)
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
        private void button3_Click(object sender, EventArgs e)
        {
          
            RmaNo frm1 = new RmaNo();
            frm1.q1 = "¶iª÷¥Í";
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    System.Data.DataTable dt1 = GetAR2(frm1.q);
                    System.Data.DataTable dt2 = rm.Rma_InvoiceDr;

                    for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                    {
                        DataRow drw = dt1.Rows[i];
                        DataRow drw2 = dt2.NewRow();
                        drw2["ShippingCode"] = shippingCodeTextBox.Text;
                        drw2["RmaNo"] = drw["U_RMA_NO"];
                        drw2["MarkNos"] = drw["U_RMODEL"];
                        drw2["SeqNo"] = "0";
                        drw2["VenderNo"] = drw["U_AUO_RMA_NO"];
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

                        drw2["Grade"] = drw["U_Rgrade"];
                        drw2["InvoiceNo_seq"] = drw["U_Rver"];
                        drw2["size"] = drw["bb"];
                        drw2["CodeName"] = drw["U_cusname_s"];
                        dt2.Rows.Add(drw2);
                    }
                    for (int j = 0; j <= rma_InvoiceDrDataGridView.Rows.Count - 2; j++)
                    {
                        rma_InvoiceDrDataGridView.Rows[j].Cells[0].Value = j.ToString();
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
        public static System.Data.DataTable GetAR2(string DocEntry)
        {
            SqlConnection MyConnection = globals.shipConnection;
            string aa = '"'.ToString();
            string sql = "select U_RMA_NO,U_AUO_RMA_NO,U_Rgrade,U_Rmodel,U_cusname_s,U_Rver,U_Rmodel,U_Rquinity,aa=case substring(U_Rmodel,4,1) when '0' then substring(U_Rmodel,2,2)+'" + aa + "'+'TFT LCD MODULE' ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1)+'" + aa + "'+'TFT LCD MODULE' END,bb=case substring(U_Rmodel,4,1) when '0' then substring(U_Rmodel,2,2) ELSE substring(U_Rmodel,2,2)+'.'+substring(U_Rmodel,4,1) END from octr where Contractid IN (" + DocEntry + ") ";
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

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {


                object[] LookupValues = GetMenu.GetMenuList();

                if (LookupValues != null)
                {
                    cardCodeTextBox.Text = Convert.ToString(LookupValues[0]);
                    cardNameTextBox.Text = Convert.ToString(LookupValues[1]);
                    shipmentTextBox.Text = Convert.ToString(LookupValues[3]);
                    add9TextBox.Text = Convert.ToString(LookupValues[5]);
                    add10TextBox.Text = Convert.ToString(LookupValues[6]);
                    add1TextBox.Text = Convert.ToString(LookupValues[2]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string aa = cardNameTextBox.Text.Substring(0, 4);
                object[] LookupValues = GetMenu.RmaCardcode(aa);

                if (LookupValues != null)
                {

                    add9TextBox.Text = Convert.ToString(LookupValues[5]);
                    add10TextBox.Text = Convert.ToString(LookupValues[6]);
                    shipmentTextBox.Text = Convert.ToString(LookupValues[3]);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                rm.Rma_InvoiceDr.RejectChanges();

            }
            catch
            {
            }

            return true;
        }
        private void Rmar_Load(object sender, EventArgs e)
        {
            System.Data.DataTable T1 = GetOHEM();

            if (T1.Rows.Count > 0)
            {
                NAME = T1.Rows[0][0].ToString();
                PHONE = T1.Rows[0][1].ToString();
                int N1 = NAME.IndexOf(" ");
                if (N1 != -1)
                {

                    NAME = NAME.Substring(0, N1);
                }


            }

            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();
            Control();

        }

        private void rma_InvoiceDrDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (rma_InvoiceDrDataGridView.Columns[e.ColumnIndex].Name == "AUOQty" ||
                    rma_InvoiceDrDataGridView.Columns[e.ColumnIndex].Name == "Qty1" ||
                    rma_InvoiceDrDataGridView.Columns[e.ColumnIndex].Name == "Qty2" ||
                 rma_InvoiceDrDataGridView.Columns[e.ColumnIndex].Name == "Qty3")
                {
                    int InQty = 0;
                    int Qty1 = 0;
                    int Qty2 = 0;
                    int Qty3 = 0;
                    if (this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["AUOQty"].Value.ToString() == "")
                    {
                        InQty = 0;
                    }
                    else
                    {

                        InQty = Convert.ToInt32(this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["AUOQty"].Value);
                    }
                    if (this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["Qty1"].Value.ToString() == "")
                    {
                        Qty1 = 0;
                    }
                    else
                    {
                        Qty1 = Convert.ToInt32(this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["Qty1"].Value);
                    }
                    if (this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["Qty2"].Value.ToString() == "")
                    {
                        Qty2 = 0;
                    }
                    else
                    {
                        Qty2 = Convert.ToInt32(this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["Qty2"].Value);
                    }
                    if (this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["Qty3"].Value.ToString() == "")
                    {
                        Qty3 = 0;
                    }
                    else
                    {
                        Qty3 = Convert.ToInt32(this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["Qty3"].Value);
                    }
                    this.rma_InvoiceDrDataGridView.Rows[e.RowIndex].Cells["Qty4"].Value = (InQty + Qty1 + Qty2 + Qty3).ToString();

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        private void rma_InvoiceDrDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = rma_InvoiceDrDataGridView.Rows.Count - 1;
            e.Row.Cells["SeqNo"].Value = iRecs.ToString();


        }


        private System.Data.DataTable GetOrderData2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("         select t0.rma_no NUM,ISNULL(tt.shippingcode,'') JOBNO,ISNULL(tt.venderno,'') VENDER,ISNULL(tt.MARKNOS,'') «~¦W³W®æ");
            sb.Append("       ,ISNULL(tt.invoiceno_seq,'') ª©¥»,ISNULL(CAST(tt.inqty AS VARCHAR),'') ¼Æ¶q,TT.¥X³f¤é´Á,TT.CODENAME+TT.RMANO RMANO,TT.»s³æ  from rma_tempr t0");
            sb.Append("              left join (select t0.shippingcode,venderno,MARKNOS,invoiceno_seq,codename,");
            sb.Append("             inqty,seqno,Convert(varchar(10),Getdate(),111) ¥X³f¤é´Á,RMANO RMANO,RANK() OVER (ORDER BY T1.doctentry DESC) §Ç¸¹,CLOSEDAY »s³æ from rma_mainr T0");
            sb.Append("              LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE T0.SHIPPINGCODE=@shippingCode  ) tt ");
            sb.Append("              on(t0.rma_no=tt.§Ç¸¹)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData4(string ½c, string RMA)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("                   select MARKNOS MODEL,INVOICENO_SEQ VER,UnitPrice QTY,A4 KG,''''+A1 PAL,''''+VenderNo VENDER,RMANO,½c=@½c,RMA=@RMA,A2 BOX,MARKNOS+'  '+'V.'+INVOICENO_SEQ «~¦W³W®æ from rma_mainr T0 ");
            sb.Append("                   LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@½c", ½c));
            command.Parameters.Add(new SqlParameter("@RMA", RMA));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData3(string ss, string PCS)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
  
            string GG = "ACME »s³æ: " + NAME + " #" + PHONE;
            string GG2 = "½ÐÃ±¦¬¦^¶Ç¦Ü¡G886-2-87912869 " + NAME + " ¦¬";

            sb.Append(" select t0.shippingcode JOBNO,RMANO RMA,''''+VenderNo VRMA,MARKNOS+'  '+'V.'+INVOICENO_SEQ «~¦W³W®æ,BRAND TOTAL,RTRIM(LTRIM(A2)) CART,RTRIM(LTRIM(A1)) PLAT,PCS=@PCS,");
            sb.Append(" auoqty AUOÁÙ¦^,QTY1 ­ì«~ÁÙ¦^,QTY2 °e­×ÁÙ¦^,QTY3 ´«³fÁÙ¦^,QTY4 ¼Æ¶q,T1.CODENAME «È¤á,T0.CARDNAME ¼t°Ó,InvoiceNo_seq VER,UnitPrice ¥X³f¼Æ¶q,");
            sb.Append(" QTY5 ÂkÁÙ,RANK() OVER (ORDER BY T1.SEQNO ) §Ç¸¹,T1.doctentry,SUBSTRING(forecastDay,1,4)+'/'+SUBSTRING(forecastDay,5,2)+'/'+SUBSTRING(forecastDay,7,2)  ¥X³f¤é´Á,T0.CLOSEDAY »s³æ,ltrim(A3) A3,ltrim(A4) A4,»s³æ¤H='" + GG + "',»s³æ¤H2='" + GG2 + "'   from rma_mainr T0");
            sb.Append(" LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE T0.SHIPPINGCODE=@shippingCode and t1.doctentry in ( " + ss.ToString() + " ) ORDER BY RANK() OVER (ORDER BY CAST(T1.SEQNO AS INT))  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PCS", PCS));



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderDataWH(string ss)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("               select RANK() OVER (ORDER BY CAST(T1.SEQNO AS INT) ) AS §Ç¸¹,t0.shippingcode JOBNO,receivePlace ­Ü®w,boatCompany ³Æµù,''''+forecastDay ¤é´Á");
            sb.Append("              ,CodeName «È¤á,RmaNo RMANO,VenderNo VENDER,MARKNOS+'  '+'V.'+INVOICENO_SEQ «~¦W³W®æ");
            sb.Append("              ,INDescription ²¾­Ü from rma_mainr T0 ");
            sb.Append("               LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE T0.SHIPPINGCODE=@shippingCode and t1.doctentry in ( " + ss.ToString() + " ) ORDER BY RANK() OVER (ORDER BY CAST(T1.SEQNO AS INT))   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetD1()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("          select ''''+VenderNo VRMA,MARKNOS MODEL,ISNULL(RTRIM(LTRIM(A2)),'') CART,ISNULL(RTRIM(LTRIM(A1)),'') PLAT,InvoiceNo_seq VER,UnitPrice QT from rma_mainr T0");
            sb.Append("          LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE T0.SHIPPINGCODE=@shippingCode ORDER BY RANK() OVER (ORDER BY CAST(T1.SEQNO AS INT))   ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData31(string FLAG)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("             select t0.shippingcode JOBNO from rma_mainr T0 ");
            if (FLAG == "1")
            {
                sb.Append("    WHERE receivePlace='¥x¥_¤º´ò' AND boatCompany ='RMA¥X³f«È¤á'  ");
            }
            else if (FLAG == "2")
            {
                sb.Append("    WHERE receivePlace¡@NOT LIKE '%¤º´ò%'¡@AND add6='§Ö»¼'  ");
            }
            sb.Append("    AND forecastDay BETWEEN @AA AND @BB ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData32()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("    select t0.shippingcode JOBNO from rma_mainr T0 WHERE receivePlace¡@NOT LIKE '%¤º´ò%'¡@AND add6='§Ö»¼' AND forecastDay BETWEEN @AA AND @BB ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOHEM()
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MOBILE,OFFICEEXT FROM OHEM WHERE HOMETEL=@HOMETEL ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@HOMETEL", fmLogin.LoginID.ToString()));





            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData31T(string FLAG)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("             select T0.ForecastDay  ¥X³f¤é´Á,t0.shippingcode JOBNO,T1.CODENAME «È¤á,T1.SEQNO+1 'NO.',RMANO «È¤áRMA,MARKNOS+'  '+'V.'+INVOICENO_SEQ «~¦W³W®æ,");
            sb.Append(" auoqty AUOÁÙ¦^,QTY1 ­ì«~ÁÙ¦^,QTY2 °e­×ÁÙ¦^,QTY3 ´«³fÁÙ¦^,QTY4 ¼Æ¶q,            ");
            sb.Append("             QTY5 ©|¥¼ÂkÁÙ¼Æ¶q from rma_mainr T0");
            sb.Append("             LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            if (FLAG == "1")
            {
                sb.Append(" WHERE receivePlace='¥x¥_¤º´ò' AND boatCompany ='RMA¥X³f«È¤á' ");
            }
            if (FLAG == "2")
            {
                sb.Append(" WHERE receivePlace NOT LIKE '%¤º´ò%'¡@AND add6='§Ö»¼' ");
            }
            sb.Append(" AND forecastDay BETWEEN @AA AND @BB  ORDER BY t0.shippingcode,T1.SEQNO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData31T2()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("                     select t0.shippingcode JOBNO,T1.CODENAME «È¤á,RMANO RMA,SUM(CAST(QTY4 AS INT)) ¼Æ¶q");
            sb.Append("              from rma_mainr T0");
            sb.Append("                        LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("            WHERE receivePlace='¥x¥_¤º´ò' AND boatCompany ='RMA¥X³f«È¤á' AND forecastDay BETWEEN '20130101' AND '20131231'");
            sb.Append("              AND ISNULL(INVOICENO_SEQ,'') <> '' AND forecastDay BETWEEN @AA AND @BB");
            sb.Append("         GROUP BY  t0.shippingcode,T1.CODENAME,RMANO");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData31T3()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("   select  max(id) ID from RMA_UPDATESAP");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;





            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData31T1()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("              select Convert(varchar(10),Getdate(),112) ¥X³f¤é´Á,t0.shippingcode JOBNO,T1.CODENAME «È¤á,RMANO «È¤áRMA,MARKNOS+'  '+'V.'+INVOICENO_SEQ «~¦W³W®æ,SUM(CAST(QTY4 AS INT)) ¼Æ¶q from rma_mainr T0");
            sb.Append("              LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append("  WHERE receivePlace='¥x¥_¤º´ò' AND boatCompany ='RMA¥X³f«È¤á' AND forecastDay BETWEEN @AA AND @BB");
            sb.Append("  AND ISNULL(MARKNOS+'  '+'V.'+INVOICENO_SEQ,'') <> '' ");
            sb.Append(" GROUP BY  t0.shippingcode,T1.CODENAME,RMANO,MARKNOS,INVOICENO_SEQ ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData31CARD(string SHIPPINGCODE)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("             select buCardname from rma_mainr T0 WHERE  SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOrderData31RMA(string SHIPPINGCODE)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("            select distinct RMANO RMA from RMA_INVOICEDR   WHERE  SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData32(string SHIPPINGCODE, string PCS)
        {

            string GG = "ACME »s³æ: " + NAME + " #" + PHONE;
            string GG2 = "½ÐÃ±¦¬¦^¶Ç¦Ü¡G886-2-87912869 " + NAME + " ¦¬";
            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select t0.shippingcode JOBNO,RMANO RMA,''''+VenderNo VRMA,MARKNOS+'  '+'V.'+INVOICENO_SEQ «~¦W³W®æ,BRAND TOTAL,RTRIM(LTRIM(A2)) CART,RTRIM(LTRIM(A1)) PLAT,PCS=@PCS,");
            sb.Append("             auoqty AUOÁÙ¦^,QTY1 ­ì«~ÁÙ¦^,QTY2 °e­×ÁÙ¦^,QTY3 ´«³fÁÙ¦^,QTY4 ¼Æ¶q,T1.CODENAME «È¤á,T0.CARDNAME ¼t°Ó,InvoiceNo_seq VER,UnitPrice ¥X³f¼Æ¶q,");
            sb.Append("             QTY5 ÂkÁÙ,RANK() OVER (ORDER BY T1.SEQNO ) §Ç¸¹,T1.doctentry,T0.ForecastDay ¥X³f¤é´Á,T0.CLOSEDAY »s³æ,ltrim(A3) A3,ltrim(A4) A4,»s³æ¤H='" + GG + "',»s³æ¤H2='" + GG2 + "'  from rma_mainr T0");
            sb.Append("             LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE  T0.SHIPPINGCODE=@SHIPPINGCODE");
            sb.Append("  ORDER BY RANK() OVER (ORDER BY T1.SEQNO) ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@PCS", PCS));



            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetOrderData3G(string ss)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select 'SAY TOTAL: '+CAST(SUM(CAST(ISNULL(A2,0) AS INT)) AS VARCHAR)+'CTN ONLY.' ¼Æ¶q,MAX(BRAND) from rma_mainr T0");
            sb.Append(" LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) WHERE T0.SHIPPINGCODE=@shippingCode and t1.doctentry in ( " + ss.ToString() + " )");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@shippingCode", shippingCodeTextBox.Text));




            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMAI()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT MARKNOS FROM rma_InvoiceD WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetCART(string MODEL_NO, string MODEL_VER)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CT_QTY,PAL_QTY,MODEL_NO,MODEL_VER FROM acmesqlSP.dbo.CART WHERE MODEL_NO=@MODEL_NO AND MODEL_VER=@MODEL_VER");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL_NO", MODEL_NO));
            command.Parameters.Add(new SqlParameter("@MODEL_VER", MODEL_VER));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetCART2(string SHI)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT MARKNOS MODEL,InvoiceNo_seq VER,SUM(CAST(ISNULL(Qty4,0) AS INT)) QTY FROM rma_mainr T0");
            sb.Append(" LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE");
            sb.Append(" GROUP BY MARKNOS,InvoiceNo_seq");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHI));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetDEL()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T0.SHIPPINGCODE JOBNO,CONVERT(VARCHAR(10) , GETDATE(), 111 ) ¥X³f¤é´Á, ");
            sb.Append(" CONVERT(VARCHAR(10), CAST(T1.DOCDATE AS DATETIME), 111 )   ¥X³f¤é´Á1,");
            sb.Append(" T0.CARDNAME,T1.CARDNAME «È¤á¦WºÙ,T1.RMANO ACMERMA,''''+T1.AURMANO AURMA,");
            sb.Append(" T1.MODEL,T1.VER,T1.QTY,T1.REPAIRCENTER REP,T1.REMARK FROM RMA_MAINR T0");
            sb.Append(" LEFT JOIN  RMA_DELIVERY T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE ORDER BY SUBSTRING(T1.MODEL,2,3) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetAUO()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT  T0.SHIPPINGCODE JOBNO,CONVERT(VARCHAR(10) , GETDATE(), 111 ) ¥X³f¤é´Á, ");
            sb.Append(" CONVERT(VARCHAR(10), CAST(T1.DOCDATE AS DATETIME), 111 )   ¥X³f¤é´Á1,");
            sb.Append(" T0.CARDNAME,T1.CARDNAME «È¤á¦WºÙ,T1.RMANO ACMERMA,T1.AURMANO AURMA,");
            sb.Append(" T1.MODEL,T1.VER,T1.QTY,T1.REPAIRCENTER REP,T1.REMARK FROM RMA_MAINR T0");
            sb.Append(" LEFT JOIN  RMA_DELIVERY T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetCART3(int AA)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME+' - '+add1 ¦¬¥ó¤H,shipment ¹q¸Ü,add10 SHIP,RMANO, MARKNOS  MODEL, Qty4  QTY, InvoiceNo_seq  VER FROM rma_mainr T0");
            sb.Append(" LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@AA", AA));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetOWTR(string U_RMA_NO)
        {


            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CONTRACTID   FROM OCTR WHERE U_RMA_NO=@U_RMA_NO ORDER BY CONTRACTID DESC");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetDELIVERY(string REPAIRCENTER)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" select RMANO,MODEL,QTY,MARK,CAST(DOCDATE AS DATETIME)  DOCDATE,''''+AURMANO AURMANO,VER,AURMANO AURMANO2  FROM RMA_DELIVERY  WHERE SHIPPINGCODE=@SHIPPINGCODE AND REPAIRCENTER=@REPAIRCENTER AND MARK > 0 AND ISNULL(DOCDATE,'') <> '' ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@REPAIRCENTER", REPAIRCENTER));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetDELIVERYAUO()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" select DISTINCT SHIPPINGCODE,''''+VENDERNO AURMANO,MARKNOS MODEL,InvoiceNo_seq VER,UnitPrice QTY,CEILING(CAST(CAST(UnitPrice AS DECIMAL(10,2))/CAST(CT_QTY AS DECIMAL(10,2)) AS DECIMAL(10,2))) MARK  FROM RMA_INVOICEDR T0");
            sb.Append(" LEFT JOIN CART T1 ON (T0.MARKNOS=T1.MODEL_NO AND T0.InvoiceNo_seq=T1.MODEL_VER)");
            sb.Append(" WHERE UnitPrice NOT IN ('100BUFFER','2 ¤ÀªR','216BUFFER') AND SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetCART31(int AA, string SHIPPINGCODE)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CARDNAME+' - '+add1 ¦¬¥ó¤H,shipment ¹q¸Ü,add10 SHIP,CASE WHEN @AA=1 THEN '' ELSE RMANO END RMANO,CASE WHEN @AA=1 THEN '' ELSE MARKNOS END MODEL,CASE WHEN @AA=1 THEN '' ELSE Qty4 END QTY,CASE WHEN @AA=1 THEN '' ELSE InvoiceNo_seq END VER FROM rma_mainr T0");
            sb.Append(" LEFT JOIN Rma_InvoiceDr T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE) ");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));
            command.Parameters.Add(new SqlParameter("@AA", AA));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetMAI2()
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT DISTINCT RMANO FROM rma_InvoiceD WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GetCARTT(string MODEL_NO, string MODEL_VER)
        {


            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT CT_QTY  FROM CART WHERE MODEL_NO=@MODEL_NO AND MODEL_VER=@MODEL_VER ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MODEL_NO", MODEL_NO));
            command.Parameters.Add(new SqlParameter("@MODEL_VER", MODEL_VER));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_mainr");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {

                if (rma_InvoiceDrDataGridView.SelectedRows.Count < 1)
                {
                    MessageBox.Show("½Ð¥ý¿ï¾Ü");
                    return;

                }

                DELETEFILE();
                string FileName = string.Empty;
                string FileName2 = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\«È¤áÃ±¦¬³æ(§Ö»¼).xls";

                FileName2 = lsAppDir + "\\Excel\\RMA\\áMÀY.xls";

                DataGridViewRow row;

                DataGridViewRow row1;
                for (int i = rma_InvoiceDrDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = rma_InvoiceDrDataGridView.SelectedRows[i];


                    listBox1.Items.Add(row.Cells["doctentry"].Value.ToString());
                    listBox2.Items.Add(row.Cells["RmaNo"].Value.ToString());
                }
                row1 = rma_InvoiceDrDataGridView.SelectedRows[0];
                string CodeName = row1.Cells["CodeName"].Value.ToString();
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    al.Add(listBox1.Items[i].ToString());
                }
                StringBuilder ss = new StringBuilder();



                foreach (string v in al)
                {
                    ss.Append("" + v + ",");
                }

                ss.Remove(ss.Length - 1, 1);


                string q = ss.ToString();
                listBox1.Items.Clear();

                ArrayList al2 = new ArrayList();

                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    al2.Add(listBox2.Items[i].ToString());
                }
                StringBuilder ss2 = new StringBuilder();

                foreach (string v in al2)
                {
                    ss2.Append("" + v + "_");
                }

                ss2.Remove(ss2.Length - 1, 1);
                string q2 = ss2.ToString();
                listBox2.Items.Clear();

                System.Data.DataTable OrderData = GetOrderData3(q, "");
                System.Data.DataTable OrderData2 = null;
                System.Data.DataTable OrderData3 = GetCART2(shippingCodeTextBox.Text);
                string MODEL = OrderData3.Rows[0]["MODEL"].ToString();
                string VER = OrderData3.Rows[0]["VER"].ToString();
                OrderData2 = GetCART3(1);
            
                    System.Data.DataTable OrderData4 = GetCART(MODEL, VER);
                    if (OrderData4.Rows.Count > 0)
                    {
                        string CC1 = OrderData4.Rows[0][0].ToString();
                        if (String.IsNullOrEmpty(CC1))
                        {
                            OrderData2 = GetCART3(1);
                        }
                        else
                        {
                            int C1 = Convert.ToInt32(CC1);
                            int C2 = Convert.ToInt32(OrderData3.Rows[0]["QTY"].ToString());
                            if (C1 >= C2)
                            {
                                OrderData2 = GetCART3(0);
                            }
                        }
                    }


                //Excelªº¼Ëª©ÀÉ
                string ExcelTemplate = FileName;
                string ExcelTemplate2 = FileName2;
                //¿é¥XÀÉ
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                    "¥X³f¤u³æ_" + CodeName + "_" + q2 + ".xls";
                string OutPutFile2 = lsAppDir + "\\Excel\\temp\\" +
    "¶iª÷¥Í±H¥XáMÀY_" + CodeName + ".xls";
                //¥X³f¤u³æ_«È¤á¦W_ RMA NO
                //²£¥Í Excel Report
                if (OrderData.Rows.Count > 0)
                {
                    ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");

                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }

                if (OrderData2.Rows.Count > 0)
                {
                    ExcelReport.ExcelReportOutput(OrderData2, ExcelTemplate2, OutPutFile2, "N");

                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }
        private void DELETEFILE()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp\\";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private void DELETEFILErmar()
        {
            try
            {
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
                string OutPutFile = lsAppDir + "\\Excel\\temp\\rmar\\";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch
            { }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            DataGridViewRow row;
            for (int i = rma_InvoiceDrDataGridView.SelectedRows.Count - 1; i >= 0; i--)
            {
                row = rma_InvoiceDrDataGridView.SelectedRows[i];

                string S = comboBox1.SelectedValue.ToString();
                string T = comboBox2.SelectedValue.ToString();
                row.Cells[T].Value = row.Cells[S].Value.ToString();
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {

                object[] LookupValues = GetMenu.RmrRRS();

                if (LookupValues != null)
                {
                    receivePlaceTextBox.Text = Convert.ToString(LookupValues[1]);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {

                object[] LookupValues = GetMenu.RmrRRM();

                if (LookupValues != null)
                {
                    boatCompanyTextBox.Text = Convert.ToString(LookupValues[1]);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {

                object[] LookupValues = GetMenu.RmrRRSH();

                if (LookupValues != null)
                {
                    buCardcodeTextBox.Text = Convert.ToString(LookupValues[1]);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {

                object[] LookupValues = GetMenu.RmrRRT();

                if (LookupValues != null)
                {
                    boatNameTextBox.Text = Convert.ToString(LookupValues[1]);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {

                object[] LookupValues = GetMenu.RmrCONS();

                if (LookupValues != null)
                {
                    buCardnameTextBox.Text = Convert.ToString(LookupValues[1]);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void UPOCTR(string U_AUO_RMA_NO, string U_YETQTY, string U_REPAIRCENTER, string CONTRACTID)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE  OCTR SET U_AUO_RMA_NO=@U_AUO_RMA_NO,U_YETQTY=@U_YETQTY,U_REPAIRCENTER=@U_REPAIRCENTER WHERE CONTRACTID=@CONTRACTID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
            command.Parameters.Add(new SqlParameter("@U_YETQTY", U_YETQTY));
            command.Parameters.Add(new SqlParameter("@U_REPAIRCENTER", U_REPAIRCENTER));
            command.Parameters.Add(new SqlParameter("@CONTRACTID", CONTRACTID));
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
        public void UPOCTR2(DateTime U_ACME_OUT, string U_AUO_RMA_NO)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE OCTR SET U_ACME_OUT=@U_ACME_OUT WHERE U_AUO_RMA_NO=@U_AUO_RMA_NO", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_ACME_OUT", U_ACME_OUT));
            command.Parameters.Add(new SqlParameter("@U_AUO_RMA_NO", U_AUO_RMA_NO));
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

        private void CalcTotals3()
        {
            try
            {


                Int32 UnitPrice = 0;


                int i = this.rma_InvoiceDrDataGridView.SelectedRows.Count - 1;
                for (int iRecs = 0; iRecs <= i; iRecs++)
                {


                    //UnitPrice
                    if (!String.IsNullOrEmpty(rma_InvoiceDrDataGridView.SelectedRows[iRecs].Cells["UnitPrice"].Value.ToString().Trim()))
                    {
                        string g = rma_InvoiceDrDataGridView.SelectedRows[iRecs].Cells["UnitPrice"].Value.ToString().Trim();
                        if (!String.IsNullOrEmpty(g))
                        {
                            UnitPrice += Convert.ToInt32(rma_InvoiceDrDataGridView.SelectedRows[iRecs].Cells["UnitPrice"].Value);

                        }
                    }
                }




                if (UnitPrice != 0)
                {
                    PCS = UnitPrice + " PCS";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private void CalcTotals4()
        {
            try
            {


                System.Data.DataTable TH = GetJINZ();

                if (TH.Rows.Count > 0)
                {
                    string f;
                    string a0 = TH.Rows[0][0].ToString();
                    int g = a0.LastIndexOf("~");
                    if (g == 0)
                    {
                        f = a0;
                    }
                    else
                    {
                        f = a0.Substring(g + 1);
                    }
                    CART = f;
                }


                if (CART != "")
                {
                    CART = "¦@" + CART + "½c";
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {

                if (rma_InvoiceDrDataGridView.SelectedRows.Count < 1)
                {
                    MessageBox.Show("½Ð¥ý¿ï¾Ü");
                    return;

                }

                DELETEFILE();
                string FileName = string.Empty;
                string FileName2 = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\«È¤áÃ±¦¬³æ(¬£¨®).xls";
                FileName2 = lsAppDir + "\\Excel\\RMA\\áMÀY.xls";


                DataGridViewRow row;

                DataGridViewRow row1;
                for (int i = rma_InvoiceDrDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = rma_InvoiceDrDataGridView.SelectedRows[i];


                    listBox1.Items.Add(row.Cells["doctentry"].Value.ToString());
                    listBox2.Items.Add(row.Cells["RmaNo"].Value.ToString());
                }
                row1 = rma_InvoiceDrDataGridView.SelectedRows[0];
                string CodeName = row1.Cells["CodeName"].Value.ToString();
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    al.Add(listBox1.Items[i].ToString());
                }
                StringBuilder ss = new StringBuilder();



                foreach (string v in al)
                {
                    ss.Append("" + v + ",");
                }

                ss.Remove(ss.Length - 1, 1);

                //linenum

                string q = ss.ToString();
                listBox1.Items.Clear();

                ArrayList al2 = new ArrayList();

                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    al2.Add(listBox2.Items[i].ToString());
                }
                StringBuilder ss2 = new StringBuilder();

                foreach (string v in al2)
                {
                    ss2.Append("" + v + "_");
                }

                ss2.Remove(ss2.Length - 1, 1);
                string q2 = ss2.ToString();
                listBox2.Items.Clear();

                // CalcTotals2();

                System.Data.DataTable OrderData = GetOrderData3(q, "");

                System.Data.DataTable OrderData2 = null;
                System.Data.DataTable OrderData3 = GetCART2(shippingCodeTextBox.Text);
                string MODEL = OrderData3.Rows[0]["MODEL"].ToString();
                string VER = OrderData3.Rows[0]["VER"].ToString();
                OrderData2 = GetCART3(1);
                //if (OrderData3.Rows.Count == 1)
                //{
                    System.Data.DataTable OrderData4 = GetCART(MODEL, VER);
                    if (OrderData4.Rows.Count > 0)
                    {
                        string CC1 = OrderData4.Rows[0][0].ToString();
                        if (String.IsNullOrEmpty(CC1))
                        {
                            OrderData2 = GetCART3(1);
                        }
                        else
                        {
                            int C1 = Convert.ToInt32(CC1);
                            int C2 = Convert.ToInt32(OrderData3.Rows[0]["QTY"].ToString());
                            if (C1 >= C2)
                            {
                                OrderData2 = GetCART3(0);
                            }
                        }
                    }
                    else
                    {
                        OrderData2 = GetCART3(0);
                    }

               // }

                string ExcelTemplate = FileName;
                string ExcelTemplate2 = FileName2;
                //¿é¥XÀÉ
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                    "¥X³f¤u³æ__" + CodeName + "_" + q2 + ".xls";


                string OutPutFile2 = lsAppDir + "\\Excel\\temp\\rmar\\" +
    "¶iª÷¥Í±H¥XáMÀY_" + CodeName + ".xls";
                //²£¥Í Excel Report
                if (OrderData.Rows.Count > 0)
                {
                    ExcelReport.ExcelReportOutputJOCELIN(OrderData, ExcelTemplate, OutPutFile, "Y", "²Ä¤GÁp-«È¤á¦¬³f¯d¦sÁp");

                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }

                if (OrderData2.Rows.Count > 0)
                {
                    ExcelReport.ExcelReportOutput(OrderData2, ExcelTemplate2, OutPutFile2, "N");

                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {

                if (rma_InvoiceDrDataGridView.SelectedRows.Count < 1)
                {
                    MessageBox.Show("½Ð¥ý¿ï¾Ü");
                    return;

                }

                DELETEFILE();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\¥X³fAUO¤u³æ(§Ö»¼).xls";



                DataGridViewRow row;

                DataGridViewRow row1;
                for (int i = rma_InvoiceDrDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = rma_InvoiceDrDataGridView.SelectedRows[i];


                    listBox1.Items.Add(row.Cells["doctentry"].Value.ToString());
                    listBox2.Items.Add(row.Cells["RmaNo"].Value.ToString());
                }
                row1 = rma_InvoiceDrDataGridView.SelectedRows[0];
                string CodeName = row1.Cells["CodeName"].Value.ToString();
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    al.Add(listBox1.Items[i].ToString());
                }
                StringBuilder ss = new StringBuilder();



                foreach (string v in al)
                {
                    ss.Append("" + v + ",");
                }

                ss.Remove(ss.Length - 1, 1);

                //linenum

                string q = ss.ToString();
                listBox1.Items.Clear();

                ArrayList al2 = new ArrayList();

                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    al2.Add(listBox2.Items[i].ToString());
                }
                StringBuilder ss2 = new StringBuilder();

                foreach (string v in al2)
                {
                    ss2.Append("" + v + "_");
                }

                ss2.Remove(ss2.Length - 1, 1);
                string q2 = ss2.ToString();
                listBox2.Items.Clear();



                CalcTotals3();



                System.Data.DataTable OrderData = GetOrderData3(q, PCS);

                //Excelªº¼Ëª©ÀÉ
                string ExcelTemplate = FileName;

                //¿é¥XÀÉ
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                    "¥X³fAUO¤u³æ(§Ö»¼)_" + CodeName + "_" + q2 + ".xls";

                //²£¥Í Excel Report
                if (OrderData.Rows.Count > 0)
                {
                    ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");

                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

            string CARDCODE = cardCodeTextBox.Text.Trim();
            int D1 = CARDCODE.IndexOf("Às¼æ");
            int D2 = CARDCODE.IndexOf("Às¬ì");
            int D3 = CARDCODE.IndexOf("¥x¤¤");
            if (D1 != -1 || D2 != -1 || D3 != -1)
            {

                string lsAppDir1 = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                string h1 = "";
                if (D1 != -1)
                {
                    h1 = "AUO(Às¼æ)³ÁÀY.xls";
                }
                if (D2 != -1)
                {
                    h1 = "AUO(Às¬ì)³ÁÀY.xls";
                }
                if (D3 != -1)
                {
                    h1 = "AUO(¥x¤¤)³ÁÀY.xls";
                }

                string FileName1 = lsAppDir1 + "\\Excel\\RMA\\" + h1.Replace("AUO", "AUT");
                GetExcelProductAUO(FileName1, h1);
            }

        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {

                if (rma_InvoiceDrDataGridView.SelectedRows.Count < 1)
                {
                    MessageBox.Show("½Ð¥ý¿ï¾Ü");
                    return;

                }

                DELETEFILE();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\¥X³fAUO¤u³æ(¬£¨®).xls";



                DataGridViewRow row;

                DataGridViewRow row1;
                for (int i = rma_InvoiceDrDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = rma_InvoiceDrDataGridView.SelectedRows[i];


                    listBox1.Items.Add(row.Cells["doctentry"].Value.ToString());
                    listBox2.Items.Add(row.Cells["RmaNo"].Value.ToString());
                }
                row1 = rma_InvoiceDrDataGridView.SelectedRows[0];
                string CodeName = row1.Cells["CodeName"].Value.ToString();
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    al.Add(listBox1.Items[i].ToString());
                }
                StringBuilder ss = new StringBuilder();



                foreach (string v in al)
                {
                    ss.Append("" + v + ",");
                }

                ss.Remove(ss.Length - 1, 1);

                //linenum

                string q = ss.ToString();
                listBox1.Items.Clear();

                ArrayList al2 = new ArrayList();

                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    al2.Add(listBox2.Items[i].ToString());
                }
                StringBuilder ss2 = new StringBuilder();

                foreach (string v in al2)
                {
                    ss2.Append("" + v + "_");
                }

                ss2.Remove(ss2.Length - 1, 1);
                string q2 = ss2.ToString();
                listBox2.Items.Clear();

                // CalcTotals2();
                CalcTotals3();

                System.Data.DataTable OrderData = GetOrderData3(q, PCS);

                //Excelªº¼Ëª©ÀÉ
                string ExcelTemplate = FileName;

                //¿é¥XÀÉ
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                    "¥X³fAUO¤u³æ(¬£¨®)_" + CodeName + "_" + q2 + ".xls";

                //²£¥Í Excel Report
                if (OrderData.Rows.Count > 0)
                {
                    ExcelReport.ExcelReportOutputJOCELIN2(OrderData, ExcelTemplate, OutPutFile, "Y", "²Ä¤GÁp-AUO ¦¬³f¦s¬dÁp");

                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

            try
            {

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                System.Data.DataTable dtCost = MakeCART();

                System.Data.DataTable T1 = GetD1();
                int N1 = 0;
                string A1 = "";
                DataRow dr = null;
                for (int i = 0; i <= T1.Rows.Count - 1; i++)
                {
                    DataRow dd = T1.Rows[i];
                    dr = dtCost.NewRow();
                    string PLAT = dd["PLAT"].ToString();

                    dr["CART"] = dd["CART"].ToString();
                    dr["VRMA"] = dd["VRMA"].ToString();
                    dr["MODEL"] = dd["MODEL"].ToString();
                    dr["VER"] = dd["VER"].ToString();
                    dr["QT"] = dd["QT"].ToString();
                    if (PLAT == "/")
                    {
                        PLAT = "";
                    }
                    if (!String.IsNullOrEmpty(PLAT) && A1 != PLAT)
                    {
                        N1 += 1;
                    }

                    if (!String.IsNullOrEmpty(PLAT))
                    {
                        A1 = PLAT;
                    }
                    dr["T1"] = N1.ToString();

                    dtCost.Rows.Add(dr);
                }

                OrderData1 = dtCost;

                string h1 = "";
                string CARDCODE = cardCodeTextBox.Text.Trim();
                int D1 = CARDCODE.IndexOf("Às¼æ");
                int D2 = CARDCODE.IndexOf("Às¬ì");
                int D3 = CARDCODE.IndexOf("¥x¤¤");
                if (D1 != -1)
                {
                    h1 = "Às¼æ³ÁÀY¬£¨®.xls";
                }
                if (D2 != -1)
                {
                    h1 = "Às¬ì³ÁÀY¬£¨®.xls";
                }
                if (D3 != -1)
                {
                    h1 = "¥x¤¤³ÁÀY¬£¨®.xls";
                }
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                               DateTime.Now.ToString("yyyyMMddHHmmss") + h1;
                string ExcelTemplate = lsAppDir + "\\Excel\\RMA\\" + h1;
                //²£¥Í Excel Report

                if (D1 != -1 || D2 != -1 || D3 != -1)
                {
                    GetExcelProduct(ExcelTemplate);
                }

                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {

                if (rma_InvoiceDrDataGridView.SelectedRows.Count < 1)
                {
                    MessageBox.Show("½Ð¥ý¿ï¾Ü");
                    return;

                }

                DELETEFILE();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\AUO¦Û¨ú¤u³æ.xls";



                DataGridViewRow row;

                DataGridViewRow row1;
                for (int i = rma_InvoiceDrDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = rma_InvoiceDrDataGridView.SelectedRows[i];


                    listBox1.Items.Add(row.Cells["doctentry"].Value.ToString());
                    listBox2.Items.Add(row.Cells["RmaNo"].Value.ToString());
                }
                row1 = rma_InvoiceDrDataGridView.SelectedRows[0];
                string CodeName = row1.Cells["CodeName"].Value.ToString();
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    al.Add(listBox1.Items[i].ToString());
                }
                StringBuilder ss = new StringBuilder();



                foreach (string v in al)
                {
                    ss.Append("" + v + ",");
                }

                ss.Remove(ss.Length - 1, 1);

                //linenum

                string q = ss.ToString();
                listBox1.Items.Clear();

                ArrayList al2 = new ArrayList();

                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    al2.Add(listBox2.Items[i].ToString());
                }
                StringBuilder ss2 = new StringBuilder();

                foreach (string v in al2)
                {
                    ss2.Append("" + v + "_");
                }

                ss2.Remove(ss2.Length - 1, 1);
                string q2 = ss2.ToString();
                listBox2.Items.Clear();

                // CalcTotals2();
                CalcTotals3();
                System.Data.DataTable OrderData = GetOrderData3(q, PCS);

                //Excelªº¼Ëª©ÀÉ
                string ExcelTemplate = FileName;

                //¿é¥XÀÉ
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                    "AUO¦Û¨ú¤u³æ_" + CodeName + "_" + q2 + ".xls";

                //²£¥Í Excel Report
                if (OrderData.Rows.Count > 0)
                {
                    ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");

                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_InvoiceDr;
            DataRow newCustomersRow = dt2.NewRow();

            int i = rma_InvoiceDrDataGridView.CurrentRow.Index;
            DataRow drw = dt2.Rows[i];

            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;
            newCustomersRow["Seqno"] = "100";
            newCustomersRow["CodeName"] = drw["CodeName"];
            newCustomersRow["A1"] = drw["A1"];
            newCustomersRow["A2"] = drw["A2"];
            newCustomersRow["RmaNo"] = drw["RmaNo"];
            newCustomersRow["MarkNos"] = drw["MarkNos"];
            newCustomersRow["InvoiceNo_seq"] = drw["InvoiceNo_seq"];
            newCustomersRow["Grade"] = drw["Grade"];
            newCustomersRow["InQty"] = drw["InQty"];
            newCustomersRow["INDescription"] = drw["INDescription"];
            newCustomersRow["Amount"] = drw["Amount"];
            newCustomersRow["UnitPrice"] = drw["UnitPrice"];

            newCustomersRow["AUOQty"] = drw["AUOQty"];
            newCustomersRow["Qty1"] = drw["Qty1"];
            newCustomersRow["Qty2"] = drw["Qty2"];
            newCustomersRow["Qty3"] = drw["Qty3"];
            newCustomersRow["Qty4"] = drw["Qty4"];
            newCustomersRow["Qty5"] = drw["Qty5"];
            newCustomersRow["VenderNo"] = drw["VenderNo"];
            newCustomersRow["A3"] = drw["A3"];
            newCustomersRow["A4"] = drw["A4"];

            try
            {
                dt2.Rows.InsertAt(newCustomersRow, rma_InvoiceDrDataGridView.Rows.Count);
                rma_InvoiceDrBindingSource.DataSource = dt2;


                for (int j = 0; j <= rma_InvoiceDrDataGridView.Rows.Count - 2; j++)
                {
                    rma_InvoiceDrDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void ´¡¤J¦CToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt2 = rm.Rma_InvoiceDr;
            DataRow newCustomersRow = dt2.NewRow();



            newCustomersRow["ShippingCode"] = shippingCodeTextBox.Text;

            newCustomersRow["SeqNo"] = 100;
            try
            {

                dt2.Rows.InsertAt(newCustomersRow, rma_InvoiceDrDataGridView.CurrentRow.Index);
                rma_InvoiceDrBindingSource.DataSource = dt2;

                for (int j = 0; j <= rma_InvoiceDrDataGridView.Rows.Count - 2; j++)
                {
                    rma_InvoiceDrDataGridView.Rows[j].Cells[0].Value = j.ToString();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            D1("1");
        }
        private void D1(string FLAG)
        {
            try
            {

                string MailContent = GlobalMailContent;
                DELETEFILErmar();
                string FileName = string.Empty;
                string FileName2 = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\«È¤áÃ±¦¬³æ(§Ö»¼).xls";

                FileName2 = lsAppDir + "\\Excel\\RMA\\áMÀY.xls";



                System.Data.DataTable OrderDataT = GetOrderData31(FLAG);

          
                if (OrderDataT.Rows.Count > 0)
                {
                    for (int j = 0; j <= OrderDataT.Rows.Count - 1; j++)
                    {
                        string SHI = OrderDataT.Rows[j][0].ToString();
                        System.Data.DataTable OrderData = GetOrderData32(SHI, "");
                        System.Data.DataTable OrderData2 = null;
                        System.Data.DataTable OrderData3 = GetCART2(SHI);
                        string MODEL = OrderData3.Rows[0]["MODEL"].ToString();
                        string VER = OrderData3.Rows[0]["VER"].ToString();
                        OrderData2 = GetCART31(1, SHI);
                        if (OrderData3.Rows.Count == 1)
                        {
                            System.Data.DataTable OrderData4 = GetCART(MODEL, VER);
                            if (OrderData4.Rows.Count > 0)
                            {
                                string G1 = OrderData4.Rows[0][0].ToString();
                                int C2 = Convert.ToInt32(OrderData3.Rows[0]["QTY"].ToString());
                                OrderData2 = GetCART31(0, SHI);
                            }

                        }
                        string CARD = GetOrderData31CARD(SHI).Rows[0][0].ToString();

                        StringBuilder sb = new StringBuilder();
                        System.Data.DataTable dtg = GetOrderData31RMA(SHI);
                        if (dtg.Rows.Count > 0)
                        {
                            for (int i = 0; i <= dtg.Rows.Count - 1; i++)
                            {

                                DataRow dd = dtg.Rows[i];


                                sb.Append(dd["RMA"].ToString() + "_");


                            }

                            sb.Remove(sb.Length - 1, 1);
                        }

                        if (FLAG == "1")
                        {
                            //Excelªº¼Ëª©ÀÉ
                            string ExcelTemplate = FileName;
                            string ExcelTemplate2 = FileName2;
                            //¿é¥XÀÉ
                            string q2 = sb.ToString();
                            string OutPutFile = lsAppDir + "\\Excel\\temp\\rmar\\" +
                                "¥X³f¤u³æ_" + CARD + "_" + q2 + ".xls";
                            string OutPutFile2 = lsAppDir + "\\Excel\\temp\\rmar\\" +
                "¶iª÷¥Í±H¥XáMÀY_" + CARD + ".xls";
                            //¥X³f¤u³æ_«È¤á¦W_ RMA NO
                            //²£¥Í Excel Report
                            if (OrderData.Rows.Count > 0)
                            {
                                ExcelReport.ExcelReportOutputJ2(OrderData, ExcelTemplate, OutPutFile);

                            }
                            else
                            {
                                MessageBox.Show("½Ð¥ý¦sÀÉ");
                            }

                            if (OrderData2.Rows.Count > 0)
                            {
                                ExcelReport.ExcelReportOutputJ2(OrderData2, ExcelTemplate2, OutPutFile2);

                            }
                            else
                            {
                                MessageBox.Show("½Ð¥ý¦sÀÉ");
                            }
                        }



                    }
                }


                System.Data.DataTable dtGetAcmeStage = GetOrderData31T(FLAG);
                int T = dtGetAcmeStage.Rows.Count;
                dataGridView1.DataSource = dtGetAcmeStage;

                GlobalMailContent = htmlMessageBody(dataGridView1).ToString();
                MailContent = GlobalMailContent;

                string template;
                StreamReader objReader;


                FileName = lsAppDir + "\\MailTemplates\\RMA2.htm";
                objReader = new StreamReader(FileName);

                template = objReader.ReadToEnd();
                objReader.Close();
                objReader.Dispose();




                template = template.Replace("##Content##", "Dear Tony");
                template = template.Replace("##Content2##", "¥H¤U¬°¤µ¤é¹w­pÂkÁÙ«È¤á²M³æ¡A½Ð¨ó§U¦w±ÆÁÙ³f¨Æ©y¡A¦p¦³¥ô¦ó°ÝÃD½Ð¨ó§U§iª¾¡AÁÂÁÂ¡C");
                template = template.Replace("##Content3##", MailContent);

                MailMessage message = new MailMessage();

                message.From = new MailAddress("workflow@acmepoint.com", "¨t²Îµo°e");
                string USER = fmLogin.LoginID.ToString() + "@acmepoint.com";
                message.To.Add(USER);



                message.Subject = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "¹w­pÂkÁÙ«È¤á²M³æ";
                message.Body = template;
                if (FLAG == "1")
                {
                    string OutPutFile1 = lsAppDir + "\\Excel\\temp\\rmar\\";
                    string[] filenames = Directory.GetFiles(OutPutFile1);
                    foreach (string file in filenames)
                    {

                        string m_File = "";

                        m_File = file;
                        data = new Attachment(m_File, MediaTypeNames.Application.Octet);

                        //ªþ¥ó†V®Æ
                        ContentDisposition disposition = data.ContentDisposition;


                        // ¥[¤J…o¥óªþ¥ó
                        message.Attachments.Add(data);

                    }
                }

                message.IsBodyHtml = true;

                SmtpClient client = new SmtpClient();
                client.Send(message);
                if (FLAG == "1")
                {
                    data.Dispose();
                }
                message.Attachments.Dispose();

                DELETEFILE();
                MessageBox.Show("±H«H¦¨¥\");

                System.Data.DataTable dtGetAcmeStage2 = GetOrderData31T2();
                if (dtGetAcmeStage2.Rows.Count > 0)
                {

                    string date = DateTime.Now.ToString("yyyyMMdd");
                    string TIME = DateTime.Now.ToString("HHmmss");
                    AddProduct(date, TIME);
                    for (int i = 0; i <= dtGetAcmeStage2.Rows.Count - 1; i++)
                    {
                        string JOBNO = dtGetAcmeStage2.Rows[i]["JOBNO"].ToString();
                        string «È¤á = dtGetAcmeStage2.Rows[i]["«È¤á"].ToString();
                        string RMA = dtGetAcmeStage2.Rows[i]["RMA"].ToString();
                        int ¼Æ¶q = Convert.ToInt32(dtGetAcmeStage2.Rows[i]["¼Æ¶q"].ToString());


                        string ID = GetOrderData31T3().Rows[0][0].ToString();
                        AddProduct2(ID, JOBNO, «È¤á, RMA, ¼Æ¶q);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

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
                strB.AppendLine("<tr><td>***  ¬dµL¸ê®Æ  ***</td></tr>");
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
                //if (iCol == 4)
                //{
                //    strB.AppendLine("<th class='style2'>" + dg.Columns[iCol].HeaderText + "</th>");
                //}
                //else if ( iCol == 5)
                //{
                //    strB.AppendLine("<th class='style3'>" + dg.Columns[iCol].HeaderText + "</th>");
                //}
                //else
                //{
                strB.AppendLine("<th>" + dg.Columns[iCol].HeaderText + "</th>");
                //}
            }
            strB.AppendLine("</tr>");

            //GridView ­n³]¦¨¤£¥i¥[¤J¤Î½s¿è¡D¡D¤£µM·|¦h¤@¦æªÅ¥Õ
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

                    //³B²zÁä­È




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
                    //if (d == 4 )
                    //{
                    //    if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                    //    {
                    //        // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                    //        strB.AppendLine("<td>&nbsp;</td>");
                    //    }
                    //    else
                    //    {

                    //        strB.AppendLine("<td class='style2'>" + dgvc.Value.ToString() + "</td>");
                    //    }
                    //}
                    //else if ( d == 5)
                    //{
                    //    if (string.IsNullOrEmpty(dgvc.Value.ToString()))
                    //    {
                    //        // strB.AppendLine("<td>&nbsp;&nbsp;&nbsp;</td>");
                    //        strB.AppendLine("<td>&nbsp;</td>");
                    //    }
                    //    else
                    //    {

                    //        strB.AppendLine("<td class='style3'>" + dgvc.Value.ToString() + "</td>");
                    //    }
                    //}
                    //else
                    //{
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
                            strB.AppendLine("<td >" + x.ToString("#,##0") + "</td>");
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
                            strB.AppendLine("<td>" + x.ToString("#,##0.00") + "</td>");
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

                            if (dg.Columns[dgvc.ColumnIndex].HeaderText.IndexOf("¤é´Á") >= 0)
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
                    //}

                }
                strB.AppendLine("</tr>");

            }
            //table footer & end of html file
            //strB.AppendLine("</table></center></body></html>");
            strB.AppendLine("</table>");
            return strB;



            //align="right"
        }



        public void AddProduct(string MAILDATE, string MAILTIME)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" INSERT INTO [AcmeSqlSP].[dbo].[RMA_UPDATESAP]");
            sb.Append("            ([MAILDATE],[MAILTIME])");
            sb.Append("      VALUES");
            sb.Append("            (@MAILDATE,@MAILTIME)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@MAILDATE", MAILDATE));
            command.Parameters.Add(new SqlParameter("@MAILTIME", MAILTIME));
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
        public void AddProduct2(string ID, string JOBNO, string CARDCODE, string CUSTRMA, int AUO1)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();

            sb.Append(" INSERT INTO [AcmeSqlSP].[dbo].[RMA_UPDATESAP1]");
            sb.Append("            ([ID],[JOBNO],[CARDCODE],[CUSTRMA],[AUO1])");
            sb.Append("      VALUES");
            sb.Append("            (@ID,@JOBNO,@CARDCODE,@CUSTRMA,@AUO1)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@JOBNO", JOBNO));
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));
            command.Parameters.Add(new SqlParameter("@CUSTRMA", CUSTRMA));
            command.Parameters.Add(new SqlParameter("@AUO1", AUO1));
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



        private void button17_Click(object sender, EventArgs e)
        {

            if (textBox3.Text == "")
            {
                MessageBox.Show("½Ð¿ï¾ÜÀÉ®×");
                return;
            }

            if (comboBox3.SelectedIndex == -1)
            {
                MessageBox.Show("½Ð¿ï¾Ü¤U©Ô­¶­±");
                return;
            }

            try
            {

                GetExcelContentGD44(textBox3.Text);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void GetExcelContentGD44(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths
            //string excelFile = Server.MapPath("~/") + @"Excel\2006.xls";
            string excelFile = ExcelFile;

            int sd1 = Convert.ToInt16(this.comboBox3.SelectedIndex.ToString()) + 1;
            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(sd1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);




            string AURMANO;
            string MODEL;
            string VER;
            string ACMERMA;
            string CARDNAME;
            string QTY;
            string DOCDATE;
            string REPAIRCENTER;
            for (int i = 2; i <= iRowCnt; i++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                AURMANO = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                range.Select();
                MODEL = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                range.Select();
                VER = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                range.Select();
                ACMERMA = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 6]);
                range.Select();
                CARDNAME = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                range.Select();
                QTY = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                range.Select();
                DOCDATE = range.Text.ToString().Trim().Replace("/", "");

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 9]);
                range.Select();
                REPAIRCENTER = range.Text.ToString().Trim();

                try
                {

                    System.Data.DataTable dt2 = rm.Rma_DELIVERY;
                    DataRow drw2 = dt2.NewRow();
                    if (!String.IsNullOrEmpty(ACMERMA))
                    {
                        drw2["DOCDATE"] = DOCDATE;
                        drw2["CARDNAME"] = CARDNAME;
                        drw2["RMANO"] = ACMERMA;
                        drw2["MODEL"] = MODEL;
                        drw2["VER"] = VER;
                        drw2["QTY"] = QTY;
                        drw2["AURMANO"] = AURMANO;
                        drw2["REPAIRCENTER"] = REPAIRCENTER;
                        drw2["shippingCode"] = shippingCodeTextBox.Text;
                        int n;
                        if (!String.IsNullOrEmpty(QTY))
                        {
                            if (int.TryParse(QTY, out n))
                            {
                                System.Data.DataTable L1 = GetCARTT(MODEL, VER);
                                if (L1.Rows.Count > 0)
                                {
                                    string CT_QTY = L1.Rows[0][0].ToString();
                                    if (int.TryParse(CT_QTY, out n))
                                    {
                                        double T1 = Convert.ToInt16(QTY);
                                        double T2 = Convert.ToInt16(CT_QTY);
                                        drw2["MARK"] = Convert.ToString(Math.Ceiling(T1 / T2));


                                    }
                                    else
                                    {
                                        drw2["MARK"] = "1";
                                    }

                                }
                                else
                                {
                                    drw2["MARK"] = "1";
                                }

                            }
                        }

                        dt2.Rows.Add(drw2);
                    }

                    for (int j = 0; j <= rma_DELIVERYDataGridView.Rows.Count - 2; j++)
                    {
                        rma_DELIVERYDataGridView.Rows[j].Cells[0].Value = j.ToString();
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
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


        private void button18_Click(object sender, EventArgs e)
        {

            for (int i = 0; i <= rma_DELIVERYDataGridView.Rows.Count - 2; i++)
            {
                DataGridViewRow row;
                row = rma_DELIVERYDataGridView.Rows[i];
                string RMANO = row.Cells["RMANO2"].Value.ToString();
                string AURMANO = row.Cells["AURMANO"].Value.ToString();
                string QTY = row.Cells["QTY"].Value.ToString();
                string REPAIRCENTER = row.Cells["REPAIRCENTER"].Value.ToString().Trim();
                System.Data.DataTable O1 = GetOWTR(RMANO);
                if (O1.Rows.Count > 0)
                {
                    string CONTRACTID = O1.Rows[0][0].ToString();
                    UPOCTR(AURMANO, QTY, REPAIRCENTER, CONTRACTID);
                }
            }
            MessageBox.Show("¶×¤J§¹¦¨");
        }

        private void button19_Click(object sender, EventArgs e)
        {
            DELETEFILE();
            if (GetDEL().Rows.Count > 0)
            {

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\¤º´ò¥X³fAU¤u³æ.xls";

                string ExcelTemplate = FileName;

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      "¤º´ò¥X³fAU¤u³æ_" + GetMenu.Day() + Path.GetFileName(FileName);

                ExcelReport.ExcelReportOutput(GetDEL(), ExcelTemplate, OutPutFile, "N");


                System.Data.DataTable T1 = GetDELIVERY("AUT(Às¬ì)");
                System.Data.DataTable T2 = GetDELIVERY("AUT(Às¼æ)");
                System.Data.DataTable T3 = GetDELIVERY("AUT(¥x¤¤)");

                if (T1.Rows.Count > 0)
                {
                    for (int i = 0; i <= T1.Rows.Count - 1; i++)
                    {
                        string AURMANO = T1.Rows[i]["AURMANO2"].ToString();
                        DateTime DOCDATE = Convert.ToDateTime(T1.Rows[i]["DOCDATE"]);
                        UPOCTR2(DOCDATE, AURMANO);
                    }

                    string FileName1 = lsAppDir + "\\Excel\\RMA\\AUT(Às¬ì)³ÁÀY.xls";
                    GetExcelProduct(FileName1, "AUT(Às¬ì)");
                }

                if (T2.Rows.Count > 0)
                {
                    for (int i = 0; i <= T2.Rows.Count - 1; i++)
                    {
                        string AURMANO = T2.Rows[i]["AURMANO2"].ToString();
                        DateTime DOCDATE = Convert.ToDateTime(T2.Rows[i]["DOCDATE"]);
                        UPOCTR2(DOCDATE, AURMANO);
                    }

                    string FileName2 = lsAppDir + "\\Excel\\RMA\\AUT(Às¼æ)³ÁÀY.xls";
                    GetExcelProduct(FileName2, "AUT(Às¼æ)");
                }
                if (T3.Rows.Count > 0)
                {
                    for (int i = 0; i <= T3.Rows.Count - 1; i++)
                    {
                        string AURMANO = T2.Rows[i]["AURMANO2"].ToString();
                        DateTime DOCDATE = Convert.ToDateTime(T2.Rows[i]["DOCDATE"]);
                        UPOCTR2(DOCDATE, AURMANO);
                    }

                    string FileName2 = lsAppDir + "\\Excel\\RMA\\AUT(¥x¤¤)³ÁÀY.xls";
                    GetExcelProduct(FileName2, "AUT(¥x¤¤)");
                }
            }
        }

        private void GetExcelProduct(string ExcelFile, string REPAIR)
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

            //¨ú±o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
            excelSheet.Name = shippingCodeTextBox.Text;

            //¨ú±o Excel ªº¨Ï¥Î°Ï°ì
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



                object Cell_From;
                object Cell_To;
                object FixCell;
                object FixCell2;
                System.Data.DataTable dtCost = MakeTableCombine();

                System.Data.DataTable T1 = GetDELIVERY(REPAIR);

                DataRow dr = null;
                int M1 = 0;
                for (int i = 0; i <= T1.Rows.Count - 1; i++)
                {
                    DataRow dd = T1.Rows[i];

                    int MARK = Convert.ToInt16(dd["MARK"].ToString());
                    for (int k = 0; k <= MARK - 1; k++)
                    {

                        dr = dtCost.NewRow();
                        dr["AURMANO"] = dd["AURMANO"].ToString();
                        dr["MODEL"] = dd["MODEL"].ToString() + " V." + dd["VER"].ToString();
                        if (MARK == 1)
                        {
                            dr["QTY"] = dd["QTY"].ToString();
                        }
                        else
                        {
                            dr["QTY"] = "";
                        }

                        M1++;
                        dtCost.Rows.Add(dr);

                        int M2 = M1 % 3;
                        if (M2 == 0)
                        {
                            dr["MARK2"] = "1";
                        }

                    }


                }


                Cell_From = "A1";
                Cell_To = "F16";
                excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);

                if (dtCost.Rows.Count > 0)
                {

                    if (dtCost.Rows.Count > 1)
                    {
                        for (int i = 0; i <= dtCost.Rows.Count - 2; i++)
                        {
                            FixCell = "A" + ((16 * (i + 1)) + 1);
                            FixCell2 = "F" + (16 * (i + 2));
                            range = excelSheet.get_Range(FixCell, FixCell2);
                            range.Select();
                            excelSheet.Paste(oMissing, oMissing);

                        }

                    }
                    for (int i = 0; i <= dtCost.Rows.Count - 1; i++)
                    {
                        string AURMANO = dtCost.Rows[i]["AURMANO"].ToString();
                        string MODEL = dtCost.Rows[i]["MODEL"].ToString();
                        string QTY = dtCost.Rows[i]["QTY"].ToString();

                        string MARK2 = dtCost.Rows[i]["MARK2"].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[6 + (i * 16), 3]);
                        range.Select();
                        range.Value2 = AURMANO;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[15 + (i * 16), 3]);
                        range.Select();
                        range.Value2 = MODEL;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[16 + (i * 16), 3]);
                        range.Select();
                        range.Value2 = QTY;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[17 + (i * 16), 9]);
                        range.Select();
                        range.Value2 = MARK2;
                    }

                    int iRowCnt2 = excelSheet.UsedRange.Cells.Rows.Count;
                    for (int iRecord = 1; iRecord <= iRowCnt2; iRecord++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                        sTemp = (string)range.Text;
                        string N1 = sTemp.Trim();
                        if (N1 == "1")
                        {
                            range.Select();
                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);
                            range.Select();
                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);
                            range.Select();
                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);
                            range.Select();
                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);
                            range.Select();
                            range.EntireRow.Insert(Microsoft.Office.Interop.Excel.XlDirection.xlDown,
                oMissing);

                            iRecord = iRecord + 5;
                        }
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 9]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

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
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }

                //¼W¥[¤@­Ó Close
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
                //¥i¥H±N Excel.exe ²M°£
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("²£¥Í¤@­ÓÀÉ®×->" + NewFileName);

                string Msg = string.Empty;

                System.Diagnostics.Process.Start(NewFileName);




            }
        }
        private void GetExcelProductAUO(string ExcelFile,string H1)
        {


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

            //¨ú±o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //  object SelectCell = "B10";
            //  Microsoft.Office.Interop.Excel.Range range = excelSheet.get_Range(SelectCell, SelectCell);
            excelSheet.Name = shippingCodeTextBox.Text;

            //¨ú±o Excel ªº¨Ï¥Î°Ï°ì
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



                object Cell_From;
                object Cell_To;
                object FixCell;
                object FixCell2;
                System.Data.DataTable dtCost = MakeTableCombine();

                System.Data.DataTable T1 = GetDELIVERYAUO();

                DataRow dr = null;
                int M1 = 0;
                for (int i = 0; i <= T1.Rows.Count - 1; i++)
                {
                    DataRow dd = T1.Rows[i];

                    int MARK = Convert.ToInt16(dd["MARK"].ToString());

                    for (int k = 0; k <= MARK - 1; k++)
                    {

                        dr = dtCost.NewRow();
                        dr["AURMANO"] = dd["AURMANO"].ToString();
                        dr["MODEL"] = dd["MODEL"].ToString() + " V." + dd["VER"].ToString();
                        if (MARK <= 1)
                        {
                            dr["QTY"] = dd["QTY"].ToString();
                        }
                        else
                        {
                            dr["QTY"] = "";
                        }

                        M1++;
                        dtCost.Rows.Add(dr);

                        int M2 = M1 % 3;
                        if (M2 == 0)
                        {
                            dr["MARK2"] = "1";
                        }

                    }


                }


                Cell_From = "A1";
                Cell_To = "F16";
                excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);

                if (dtCost.Rows.Count > 0)
                {

                    if (dtCost.Rows.Count > 1)
                    {
                        for (int i = 0; i <= dtCost.Rows.Count - 2; i++)
                        {
                            FixCell = "A" + ((16 * (i + 1)) + 1);
                            FixCell2 = "F" + (16 * (i + 2));
                            range = excelSheet.get_Range(FixCell, FixCell2);
                            range.Select();
                            excelSheet.Paste(oMissing, oMissing);

                        }

                    }
                    for (int i = 0; i <= dtCost.Rows.Count - 1; i++)
                    {
                        string AURMANO = dtCost.Rows[i]["AURMANO"].ToString();
                        string MODEL = dtCost.Rows[i]["MODEL"].ToString();
                        string QTY = dtCost.Rows[i]["QTY"].ToString();

                        string MARK2 = dtCost.Rows[i]["MARK2"].ToString();
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[6 + (i * 16), 3]);
                        range.Select();
                        range.Value2 = AURMANO;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[7 + (i * 16), 3]);
                        range.Select();
                        range.Value2 = add1TextBox.Text + " RMA Receiving";

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[9 + (i * 16), 2]);
                        range.Select();
                        range.Value2 = "TEL: " + shipmentTextBox.Text;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[15 + (i * 16), 3]);
                        range.Select();
                        range.Value2 = MODEL;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[16 + (i * 16), 3]);
                        range.Select();
                        range.Value2 = QTY;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.Cells[17 + (i * 16), 9]);
                        range.Select();
                        range.Value2 = MARK2;
                    }


                   

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 9]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                }
            }
            finally
            {


                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + H1;

                try
                {


                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //¼W¥[¤@­Ó Close
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
                //¥i¥H±N Excel.exe ²M°£
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("²£¥Í¤@­ÓÀÉ®×->" + NewFileName);

                string Msg = string.Empty;

                System.Diagnostics.Process.Start(NewFileName);




            }
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("AURMANO", typeof(string));
            dt.Columns.Add("MODEL", typeof(string));
            dt.Columns.Add("QTY", typeof(string));
            dt.Columns.Add("MARK", typeof(string));
            dt.Columns.Add("MARK2", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeCART()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("CART", typeof(string));
            dt.Columns.Add("VRMA", typeof(string));
            dt.Columns.Add("MODEL", typeof(string));
            dt.Columns.Add("VER", typeof(string));
            dt.Columns.Add("QT", typeof(string));
            dt.Columns.Add("T1", typeof(string));

            return dt;
        }
        private void button20_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                comboBox3.SelectedIndex = -1;
                comboBox3.Items.Clear();
                FileName = openFileDialog1.FileName;
                this.textBox3.Text = FileName;
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false;

                //Interop params
                object oMissing = System.Reflection.Missing.Value;

                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(FileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                string Count_Sheet = excelBook.Sheets.Count.ToString();
                int i = excelBook.Sheets.Count;

                for (int xi = 1; xi <= i; xi++)
                {

                    Microsoft.Office.Interop.Excel.Worksheet excelsheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(xi);
                    string X1 = xi.ToString();
                    string X2 = excelsheet.Name.ToString();
                    string name_sheet = X1 + ":" + X2;
                    comboBox3.Items.Add(name_sheet);

                }
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


                excelApp = null;
                excelBook = null;

                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();


            }
        }


        private void GetExcelProduct(string ExcelFile)
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

            //¨ú±o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);



            //¨ú±o Excel ªº¨Ï¥Î°Ï°ì
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
                //    int DetailRow1 = 0;
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

                        //ÀË¬d¬O¤£¬O Detail Row
                        //­n¥ý§@§¹©Ò¦³ Master ¤§«á¦A¥h§@ Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            // DetailRow1 = 9;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData1.Rows.Count - 1; aRow++)
                    {

                        //³Ì«á¤@µ§¤£§@
                        if (aRow != OrderData1.Rows.Count - 1)
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

                int iRowCnt3 = OrderData1.Rows.Count;
                string numString = GetOrderData3CNO().Rows[0][0].ToString();

                int number1 = 0;
                bool canConvert = int.TryParse(numString, out number1);
                if (canConvert == true && numString != "0")
                {
                    Cell_From = "A1";
                    Cell_To = "J" + Convert.ToString(iRowCnt3 + 17);

                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);
                    range.Select();
                    for (int aRow = 3; aRow <= iRowCnt3 + 17; aRow++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                        range.Select();
                        string N1 = "1";
                        range.Value2 = N1.ToString();
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                    range.Select();
                    range.Value2 = GetOrderData3CNO1().Rows[0][0].ToString();

                    int F1 = Convert.ToInt16(numString);
                    for (int i = 1; i <= F1 - 1; i++)
                    {

                        int COPY = ((iRowCnt3 + 17) * i) + 1;

                        int COPY2 = ((iRowCnt3 + 17) * (i + 1)) + 1;
                        FixCell = "A" + Convert.ToString(COPY);
                        FixCell2 = "J" + Convert.ToString(COPY2);
                        range = excelSheet.get_Range(FixCell, FixCell2);
                        range.Select();
                        excelSheet.Paste(oMissing, oMissing);

                        for (int aRow = COPY; aRow <= COPY2; aRow++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                            range.Select();
                            string N2 = (i + 1).ToString();
                            range.Value2 = N2.ToString();

                        }

                        int COPY3 = ((iRowCnt3 + 17) * i) + 5;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[COPY3, 3]);
                        range.Select();
                        range.Value2 = GetOrderData3CNO1().Rows[i][0].ToString();
                    }

                    int iRowCnt2 = ((iRowCnt3 + 17) * (F1)) + 1;

                    string FLAG = "";
                    for (int aRow = 1; aRow <= iRowCnt2; aRow++)
                    {
                        if (FLAG == "Y")
                        {

                            aRow = aRow - 1;
                            FLAG = "";
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 9]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID1 = sTemp;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
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


                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);


                try
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 9]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 9]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //¼W¥[¤@­Ó Close
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
                //¥i¥H±N Excel.exe ²M°£
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("²£¥Í¤@­ÓÀÉ®×->" + NewFileName);

                string Msg = string.Empty;

                System.Diagnostics.Process.Start(NewFileName);




            }
        }

        private void GetExcelProductJINJI(string ExcelFile)
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

            //¨ú±o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);



            //¨ú±o Excel ªº¨Ï¥Î°Ï°ì
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
                //    int DetailRow1 = 0;
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

                        //ÀË¬d¬O¤£¬O Detail Row
                        //­n¥ý§@§¹©Ò¦³ Master ¤§«á¦A¥h§@ Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            // DetailRow1 = 9;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData1.Rows.Count - 1; aRow++)
                    {

                        //³Ì«á¤@µ§¤£§@
                        if (aRow != OrderData1.Rows.Count - 1)
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

                int iRowCnt3 = OrderData1.Rows.Count;
                string numString = GetOrderData3CNOJINZI().Rows[0][0].ToString();

                int number1 = 0;
                bool canConvert = int.TryParse(numString, out number1);
                if (canConvert == true)
                {
                    Cell_From = "A1";
                    Cell_To = "J" + Convert.ToString(iRowCnt3 + 17);

                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);
                    range.Select();
                    for (int aRow = 3; aRow <= iRowCnt3 + 17; aRow++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                        range.Select();
                        string N1 = "1";
                        range.Value2 = N1.ToString();
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 3]);
                    range.Select();
                    range.Value2 = GetOrderData3CNO1().Rows[0][0].ToString();

                    int F1 = Convert.ToInt16(numString);
                    for (int i = 1; i <= F1 - 1; i++)
                    {

                        int COPY = ((iRowCnt3 + 17) * i) + 1;

                        int COPY2 = ((iRowCnt3 + 17) * (i + 1)) + 1;
                        FixCell = "A" + Convert.ToString(COPY);
                        FixCell2 = "J" + Convert.ToString(COPY2);
                        range = excelSheet.get_Range(FixCell, FixCell2);
                        range.Select();
                        excelSheet.Paste(oMissing, oMissing);

                        for (int aRow = COPY; aRow <= COPY2; aRow++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                            range.Select();
                            string N2 = (i + 1).ToString();
                            range.Value2 = N2.ToString();

                        }

                        int COPY3 = ((iRowCnt3 + 17) * i) + 5;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[COPY3, 3]);
                        range.Select();
                        range.Value2 = GetOrderData3CNO1().Rows[i][0].ToString();
                    }

                    int iRowCnt2 = ((iRowCnt3 + 17) * (F1)) + 1;

                    string FLAG = "";
                    for (int aRow = 1; aRow <= iRowCnt2; aRow++)
                    {
                        if (FLAG == "Y")
                        {

                            aRow = aRow - 1;
                            FLAG = "";
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 9]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID1 = sTemp;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
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


                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);


                try
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 9]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 9]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //¼W¥[¤@­Ó Close
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
                //¥i¥H±N Excel.exe ²M°£
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("²£¥Í¤@­ÓÀÉ®×->" + NewFileName);

                string Msg = string.Empty;

                System.Diagnostics.Process.Start(NewFileName);




            }
        }
        private void GetExcelProductJINJI2(string ExcelFile)
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

            //¨ú±o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);



            //¨ú±o Excel ªº¨Ï¥Î°Ï°ì
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
                //    int DetailRow1 = 0;
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

                        //ÀË¬d¬O¤£¬O Detail Row
                        //­n¥ý§@§¹©Ò¦³ Master ¤§«á¦A¥h§@ Detail
                        if (IsDetailRow(sTemp))
                        {
                            IsDetail = true;
                            DetailRow = iRecord;
                            // DetailRow1 = 9;
                            break;
                        }


                    }

                }

                if (DetailRow != 0)
                {

                    for (int aRow = 0; aRow <= OrderData1.Rows.Count - 1; aRow++)
                    {

                        //³Ì«á¤@µ§¤£§@
                        if (aRow != OrderData1.Rows.Count - 1)
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

                int iRowCnt3 = OrderData1.Rows.Count;
                string numString = GetOrderData3CNOJINZI2().Rows[0][0].ToString();

                int number1 = 0;
                bool canConvert = int.TryParse(numString, out number1);
                if (canConvert == true)
                {
                    Cell_From = "A1";
                    Cell_To = "J" + Convert.ToString(iRowCnt3 + 17);

                    excelSheet.get_Range(Cell_From, Cell_To).Copy(oMissing);
                    range.Select();
                    for (int aRow = 3; aRow <= iRowCnt3 + 17; aRow++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                        range.Select();
                        string N1 = "1";
                        range.Value2 = N1.ToString();
                    }

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[5, 2]);
                    range.Select();
                    range.Value2 = "";



                    int F1 = Convert.ToInt16(numString);
                    for (int i = 1; i <= F1 - 1; i++)
                    {

                        int COPY = ((iRowCnt3 + 17) * i) + 1;

                        int COPY2 = ((iRowCnt3 + 17) * (i + 1)) + 1;
                        FixCell = "A" + Convert.ToString(COPY);
                        FixCell2 = "J" + Convert.ToString(COPY2);
                        range = excelSheet.get_Range(FixCell, FixCell2);
                        range.Select();
                        excelSheet.Paste(oMissing, oMissing);

                        for (int aRow = COPY; aRow <= COPY2; aRow++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
                            range.Select();
                            string N2 = (i + 1).ToString();
                            range.Value2 = N2.ToString();

                        }

                        int COPY3 = ((iRowCnt3 + 17) * i) + 5;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[COPY3, 3]);
                        range.Select();
                        range.Value2 = GetOrderData3CNO12().Rows[i][0].ToString();
                    }

                    int iRowCnt2 = ((iRowCnt3 + 17) * (F1)) + 1;

                    string FLAG = "";
                    for (int aRow = 1; aRow <= iRowCnt2; aRow++)
                    {
                        if (FLAG == "Y")
                        {

                            aRow = aRow - 1;
                            FLAG = "";
                        }

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 9]);
                        range.Select();
                        sTemp = (string)range.Text;
                        sTemp = sTemp.Trim();
                        ID1 = sTemp;

                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[aRow, 10]);
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


                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(ExcelFile);


                try
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 9]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, 9]);
                    range.Select();
                    range.EntireColumn.Delete(XlDirection.xlToLeft);

                    excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
                }

                //¼W¥[¤@­Ó Close
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
                //¥i¥H±N Excel.exe ²M°£
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("²£¥Í¤@­ÓÀÉ®×->" + NewFileName);

                string Msg = string.Empty;

                System.Diagnostics.Process.Start(NewFileName);




            }
        }
        private System.Data.DataTable GetOrderData3CNO()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT COUNT(*)  FROM rma_InvoiceDr WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(A1,'') <> ''  ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


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
        private System.Data.DataTable GetOrderData3CNOJINZI()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT SUBSTRING(A1,CHARINDEX('/', A1)+1,3) FROM rma_InvoiceDr WHERE SHIPPINGCODE=@SHIPPINGCODE AND CHARINDEX('/', A1) <> 0 ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


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

        private System.Data.DataTable GetOrderData3CNOJINZI2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT top 1 A2 FROM rma_InvoiceDr WHERE SHIPPINGCODE=@SHIPPINGCODE  and isnull(A2,'') <> ''  order by seqno desc");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


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
        private System.Data.DataTable GetOrderData3CNO1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT ''''+A1  FROM rma_InvoiceDr WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(A1,'') <> ''  ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


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

        private System.Data.DataTable GetOrderData3CNO12()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT ''''+A2  FROM rma_InvoiceDr WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(A2,'') <> ''  ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


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
        private System.Data.DataTable GetJINZ()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT TOP 1 A2 FROM Rma_InvoiceDr  WHERE SHIPPINGCODE=@SHIPPINGCODE  AND ISNULL(A2,'') <> ''  ORDER BY SEQNO DESC  ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


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

        private System.Data.DataTable GetJINZ2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT CODENAME FROM Rma_InvoiceDr  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(CODENAME,'') <> ''   ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


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


        private System.Data.DataTable GetJINZ3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();


            sb.Append(" SELECT DISTINCT RMANO FROM Rma_InvoiceDr  WHERE SHIPPINGCODE=@SHIPPINGCODE AND ISNULL(RMANO,'') <> ''   ");



            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));


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
                //Master ©T©w²Ä¤@µ§
                FieldValue = Convert.ToString(OrderData1.Rows[iRow][FieldName]);
            }

        }
        private void SetRow2(int iRow, string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return;
            }
            if (sData.Substring(0, 2) == "[[")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master ©T©w²Ä¤@µ§
                FieldValue = Convert.ToString(OrderData2.Rows[iRow][FieldName]);
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
                //Master ©T©w²Ä¤@µ§
                FieldValue = Convert.ToString(OrderData1.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }
        private bool CheckSerial2(string sData, ref string FieldValue)
        {
            string FieldName = string.Empty;

            if (sData.Length < 2)
            {
                return false;
            }
            if (sData.Substring(0, 2) == "<<")
            {
                FieldName = sData.Substring(2, sData.Length - 4);
                //Master ©T©w²Ä¤@µ§
                FieldValue = Convert.ToString(OrderData2.Rows[0][FieldName]);
                return true;
            }
            //}
            return false;
        }
        private void button21_Click(object sender, EventArgs e)
        {
            try
            {


                DELETEFILE();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\´º´¼¥X³f¤å¥ó.xls";

                StringBuilder sb = new StringBuilder();
                System.Data.DataTable dt = GetJINZ2();
                string CODENAME = "";
                if (dt.Rows.Count > 0)
                {
                    for (int i = 0; i <= dt.Rows.Count - 1; i++)
                    {

                        DataRow dd = dt.Rows[i];


                        sb.Append(dd["CODENAME"].ToString() + "_");


                    }

                    sb.Remove(sb.Length - 1, 1);

                    CODENAME = sb.ToString();
                }
                StringBuilder sb2 = new StringBuilder();
                System.Data.DataTable dt2 = GetJINZ3();
                string RMANO = "";
                if (dt2.Rows.Count > 0)
                {
                    for (int i = 0; i <= dt2.Rows.Count - 1; i++)
                    {

                        DataRow dd = dt2.Rows[i];


                        sb2.Append(dd["RMANO"].ToString() + "_");


                    }

                    sb2.Remove(sb2.Length - 1, 1);

                    RMANO = sb2.ToString();
                }

                CalcTotals4();

                System.Data.DataTable OrderData = GetOrderData4(CART, RMANO);

                //Excelªº¼Ëª©ÀÉ
                string ExcelTemplate = FileName;


                //²£¥Í Excel Report
                if (OrderData.Rows.Count > 0)
                {
                    string CARTH = "¬£¨®";
                    if (GetOrderData3CNOJINZI().Rows.Count > 0)
                    {

                        CARTH = "¬£¨®";


                    }
                    else if (GetOrderData3CNOJINZI2().Rows.Count > 0)
                    {

                        CARTH = "§Ö»¼";
                    }

                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                        "¥X³f´º´¼¤u³æ(" + CARTH + ")_" + CODENAME + "_" + RMANO + ".xls";


                    ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }

            if ((add6ComboBox.Text == "±M¨®" || add6ComboBox.Text == "§Ö»¼") && tradeConditionComboBox.Text == "±H¥ó")
            {


                try
                {

                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


                    System.Data.DataTable dtCost = MakeCART();

                    System.Data.DataTable T1 = GetD1();
                    int N1 = 0;
                    string A1 = "";
                    DataRow dr = null;
                    for (int i = 0; i <= T1.Rows.Count - 1; i++)
                    {
                        DataRow dd = T1.Rows[i];
                        dr = dtCost.NewRow();
                        string PLAT = dd["PLAT"].ToString();
                        string CART = dd["CART"].ToString();
                        dr["CART"] = dd["CART"].ToString();
                        dr["VRMA"] = dd["VRMA"].ToString();
                        dr["MODEL"] = dd["MODEL"].ToString();
                        dr["VER"] = dd["VER"].ToString();
                        dr["QT"] = dd["QT"].ToString();
                        if (PLAT == "/")
                        {
                            PLAT = "";
                        }
                        if (GetOrderData3CNOJINZI().Rows.Count > 0)
                        {
                            if (!String.IsNullOrEmpty(PLAT) && A1 != PLAT)
                            {
                                N1 += 1;
                            }
                        }
                        else if (GetOrderData3CNOJINZI2().Rows.Count > 0)
                        {
                            if (!String.IsNullOrEmpty(CART) && A1 != CART)
                            {
                                N1 += 1;
                            }
                        }
                        if (!String.IsNullOrEmpty(PLAT))
                        {
                            A1 = PLAT;
                        }
                        dr["T1"] = N1.ToString();

                        dtCost.Rows.Add(dr);
                    }

                    OrderData1 = dtCost;


                    string h1 = "´º´¼³ÁÀY.xls";

                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                                   DateTime.Now.ToString("yyyyMMddHHmmss") + h1;
                    string ExcelTemplate = lsAppDir + "\\Excel\\RMA\\" + h1;
                    //²£¥Í Excel Report
                    if (GetOrderData3CNOJINZI().Rows.Count > 0)
                    {

                        GetExcelProductJINJI(ExcelTemplate);


                    }
                    else if (GetOrderData3CNOJINZI2().Rows.Count > 0)
                    {

                        GetExcelProductJINJI2(ExcelTemplate);


                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);

                }
            }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            try
            {

                if (rma_InvoiceDrDataGridView.SelectedRows.Count < 1)
                {
                    MessageBox.Show("½Ð¥ý¿ï¾Ü");
                    return;

                }

                DELETEFILE();
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\©ñ³f³æ.xls";



                DataGridViewRow row;

                DataGridViewRow row1;
                for (int i = rma_InvoiceDrDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {

                    row = rma_InvoiceDrDataGridView.SelectedRows[i];


                    listBox1.Items.Add(row.Cells["doctentry"].Value.ToString());
                    listBox2.Items.Add(row.Cells["RmaNo"].Value.ToString());
                }
                row1 = rma_InvoiceDrDataGridView.SelectedRows[0];
                string CodeName = row1.Cells["CodeName"].Value.ToString();
                ArrayList al = new ArrayList();

                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    al.Add(listBox1.Items[i].ToString());
                }
                StringBuilder ss = new StringBuilder();



                foreach (string v in al)
                {
                    ss.Append("" + v + ",");
                }

                ss.Remove(ss.Length - 1, 1);

                //linenum

                string q = ss.ToString();
                listBox1.Items.Clear();

                ArrayList al2 = new ArrayList();

                for (int i = 0; i <= listBox2.Items.Count - 1; i++)
                {
                    al2.Add(listBox2.Items[i].ToString());
                }
                StringBuilder ss2 = new StringBuilder();

                foreach (string v in al2)
                {
                    ss2.Append("" + v + "_");
                }

                ss2.Remove(ss2.Length - 1, 1);
                string q2 = ss2.ToString();
                listBox2.Items.Clear();

                // CalcTotals2();
                CalcTotals3();
                System.Data.DataTable OrderData = GetOrderDataWH(q);

                //Excelªº¼Ëª©ÀÉ
                string ExcelTemplate = FileName;

                //¿é¥XÀÉ
                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                    "©ñ³f³æ_" + q2 + ".xls";

                //²£¥Í Excel Report
                if (OrderData.Rows.Count > 0)
                {
                    ExcelReport.ExcelReportOutputRMAWH(OrderData, ExcelTemplate, OutPutFile, "Y");

                }
                else
                {
                    MessageBox.Show("½Ð¥ý¦sÀÉ");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void button23_Click(object sender, EventArgs e)
        {

            D1("2");
        }


    }
    }



using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections;
using System.Reflection;
using System.IO;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
namespace ACME
{
    public partial class DRSTT : ACME.fmBase1
    {
        private decimal sd;
        string daXX;
        System.Data.DataTable LCXX;
        System.Data.DataTable LCXX2;
        public DRSTT()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            sATTTableAdapter.Connection = MyConnection;
            sATT1TableAdapter.Connection = MyConnection;
            sATT2TableAdapter.Connection = MyConnection;


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
                sa.SATT1.RejectChanges();
                sa.SATT2.RejectChanges();
          
            }
            catch
            {
            }
            return true;
        }
        public override void AfterCancelEdit()
        {
            Control();
        }

        public override void query()
        {
            button3.Enabled = true;
            tTDateTextBox.ReadOnly = false;

        }
        private void Control()
        {

            tTDateTextBox.ReadOnly = true;
            button1.Enabled = true;
            button3.Enabled = true;
            button6.Enabled = true;
            button5.Enabled = true;
            button20.Enabled = true;
            textBox2.ReadOnly = false;
            textBox3.ReadOnly = false;
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox7.ReadOnly = false;
            textBox8.ReadOnly = false;
            textBox10.ReadOnly = false;
            textBox11.ReadOnly = false;
            button7.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = true;
            button10.Enabled = true;
            button11.Enabled = true;
            button12.Enabled = true;

            button20.Enabled = true;


            button22.Enabled = true;

            checkBox1.Enabled = true;
            checkBox2.Enabled = true;
            checkBox3.Enabled = true;
            checkBox4.Enabled = true;
        }

        public override void AfterEdit()
        {
            tTDateTextBox.ReadOnly = true;
        }

        public override void AfterEndEdit()
        {
            try
            {
                System.Data.DataTable dt1 = GetTT(tTCodeTextBox.Text);
                UpdateTT1(tTCodeTextBox.Text);
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {

                    DataRow row = dt1.Rows[i];
                    string id = row["id"].ToString();
                    decimal USDAMOUNT = Convert.ToDecimal(row["USDAMOUNT"]);
                    decimal NTDAMOUNT = Convert.ToDecimal(row["NTDAMOUNT"]);
                    System.Data.DataTable dt2 = GetTT2(id, tTCodeTextBox.Text);
                    if (dt2.Rows.Count > 0)
                    {

                        DataRow row2 = dt2.Rows[0];
                        string Currency = row2["Currency"].ToString();

                        if (Currency == "USD")
                        {
                            UpdateTTUSD(USDAMOUNT, id, tTCodeTextBox.Text);
                        }

                        if (Currency == "NTD")
                        {
                            UpdateTTUSD(NTDAMOUNT, id, tTCodeTextBox.Text);
                        }

                        if (Currency == "RMB")
                        {
                            UpdateTTRMB(NTDAMOUNT, id, tTCodeTextBox.Text);
                        }
                    }

                }
                sATT1TableAdapter.Fill(sa.SATT1, MyID);


     
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public override void SetDefaultValue()
        {

            string NumberName = "TT" + DateTime.Now.ToString("yyyy");
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);

            this.tTCodeTextBox.Text = NumberName + AutoNum;
            tTDateTextBox.Text = DateTime.Now.ToString("yyyyMMdd");

            this.sATTBindingSource.EndEdit();
        }
        public override void SetInit()
        {

            MyBS = sATTBindingSource;
            MyTableName = "SATT";
            MyIDFieldName = "TTCode";
        }
        public override void FillData()
        {

            try
            {
                sATTTableAdapter.Fill(sa.SATT, MyID);
                sATT1TableAdapter.Fill(sa.SATT1, MyID);
                sATT2TableAdapter.Fill(sa.SATT2, MyID);


                decimal iTotal = 0;

                try
                {


                    int i = this.sATT1DataGridView.Rows.Count - 1;
                    for (int iRecs = 0; iRecs <= i; iRecs++)
                    {

                        iTotal += Convert.ToDecimal(sATT1DataGridView.Rows[iRecs].Cells["NTD2"].Value);



                    }
                }
                catch (Exception ex)
                {
                }
                label16.Text = "?????x?????B " + iTotal.ToString("#,##0"); ;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public SqlDataAdapter GetAdapter(object tableAdapter)
        {

            Type tableAdapterType = tableAdapter.GetType();

            SqlDataAdapter adapter = (SqlDataAdapter)tableAdapterType.GetProperty("Adapter", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(tableAdapter, null);

            return adapter;

        }
        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {

                sATTTableAdapter.Connection.Open();



                Validate();

                sATTBindingSource.EndEdit();
                sATT1BindingSource.EndEdit();
                sATT2BindingSource.EndEdit();


                ///?`?N: 4. ???? Transaction

                tx = sATTTableAdapter.Connection.BeginTransaction();



                SqlDataAdapter oWhsAdapter = GetAdapter(sATTTableAdapter);
                oWhsAdapter.UpdateCommand.Transaction = tx;
                oWhsAdapter.InsertCommand.Transaction = tx;
                oWhsAdapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter oWhsAdapter1 = GetAdapter(sATT1TableAdapter);
                oWhsAdapter1.UpdateCommand.Transaction = tx;
                oWhsAdapter1.InsertCommand.Transaction = tx;
                oWhsAdapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter oWhsAdapter2 = GetAdapter(sATT2TableAdapter);
                oWhsAdapter2.UpdateCommand.Transaction = tx;
                oWhsAdapter2.InsertCommand.Transaction = tx;
                oWhsAdapter2.DeleteCommand.Transaction = tx;

        


                sATTTableAdapter.Update(sa.SATT);
                sATT1TableAdapter.Update(sa.SATT1);
                sATT2TableAdapter.Update(sa.SATT2);
   

                this.MyID = this.tTCodeTextBox.Text;
                tx.Commit();


                UpdateData = true;
            }
            catch (Exception ex)
            {
                if (tx != null)
                {

                    tx.Rollback();

                }

                MessageBox.Show(ex.Message, "???s???~", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.sATTTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
        private void UpdateTTUSD(decimal TTUSD, string id, string ttcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update SATT1 set tttotal=TotalAmount-@TTUSD,Detail='?w????' where seqno=@id and ttcode=@ttcode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@TTUSD", TTUSD));
            command.Parameters.Add(new SqlParameter("@id", id));
            command.Parameters.Add(new SqlParameter("@ttcode", ttcode));
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

        private void UpdateTTRMB(decimal TTUSD, string id, string ttcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update SATT1 set tttotal=NTD-@TTUSD,Detail='?w????' where seqno=@id and ttcode=@ttcode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@TTUSD", TTUSD));
            command.Parameters.Add(new SqlParameter("@id", id));
            command.Parameters.Add(new SqlParameter("@ttcode", ttcode));
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

        private void Updatepath(string filename, string path, string TTCode, string ID)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update SATT1 set filename=@filename,[path]=@path where TTCode=@TTCode and ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@filename", filename));
            command.Parameters.Add(new SqlParameter("@path", path));
            command.Parameters.Add(new SqlParameter("@TTCode", TTCode));
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
        private void UpdateTT1(string ttcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE  satt1 SET TTTOTAL=NULL,Detail='' where tTcOde=@ttcode AND SEQNO IN (SELECT distinct seqno FROM SATT1 where ttcode=@ttcode and seqno not in (select distinct id from satt2 where ttcode=@ttcode))");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@ttcode", ttcode));


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
        public static System.Data.DataTable GETOINV2()
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.Connection;
            sb.Append(" SELECT ID1 ID,CARDCODE2 CARDCODE    FROM sATT2 T0");
            sb.Append(" INNER JOIN AcmeSql02.DBO.OCRD T1 ON (T0.CARDCODE2=T1.CARDCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append("  WHERE ISNULL(T0.CARDFNAME,'')=''  AND ISNULL(T1.CARDFNAME,'') <>'' ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }

        public static System.Data.DataTable GETOINV3()
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.Connection;
            sb.Append(" SELECT ID1 ID,DOCENTRY    FROM sATT2 WHERE ISNULL(SA,'')='' and docentry > 36443 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }
        public static System.Data.DataTable GETOINV(string CARDCODE)
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.shipConnection;
            sb.Append("SELECT CARDFNAME FROM OCRD WHERE CARDCODE=@CARDCODE AND ISNULL(CARDFNAME,'') <> '' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CARDCODE", CARDCODE));

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }
   
        public void OINV()
        {


            System.Data.DataTable G2 = GETOINV2();
            if (G2.Rows.Count > 0)
            {
                for (int i = 0; i <= G2.Rows.Count - 1; i++)
                {
                    string ID = G2.Rows[i]["ID"].ToString();
                    string CARDCODE = G2.Rows[i]["CARDCODE"].ToString();

                    System.Data.DataTable G1 = GETOINV(CARDCODE);
                    if (G1.Rows.Count > 0)
                    {
                        UpdateOINV(ID, G1.Rows[0][0].ToString());
                    }

           
                }

            }

            
            System.Data.DataTable G4 = GETOINV3();
            if (G4.Rows.Count > 0)
            {
                for (int i = 0; i <= G4.Rows.Count - 1; i++)
                {
                    string ID = G4.Rows[i]["ID"].ToString();
                    string DOCENTRY = G4.Rows[i]["DOCENTRY"].ToString();
                    System.Data.DataTable G3 = GetMenu.GetSA(DOCENTRY);
                    if (G3.Rows.Count > 0)
                    {
                        string SA = G3.Rows[0]["?~??"].ToString();
                        string SALES = G3.Rows[0]["?~??"].ToString();

                        UpdateOINV2(ID, SA, SALES);
                    }
                }

            }

        }
        private void UpdateOINV(string ID1, string CARDFNAME)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE SATT2 SET CARDFNAME =@CARDFNAME WHERE ID1=@ID1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@CARDFNAME", CARDFNAME));
            command.Parameters.Add(new SqlParameter("@ID1", ID1));

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

        private void UpdateOINV2(string ID1, string SA, string SALES)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE SATT2 SET SA =@SA,SALES=@SALES WHERE ID1=@ID1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@SA", SA));
            command.Parameters.Add(new SqlParameter("@SALES", SALES));
            command.Parameters.Add(new SqlParameter("@ID1", ID1));

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
        private void TT_Load(object sender, EventArgs e)
        {

            Control();
            OINV();


            label1.Text = "";
            label2.Text = "";
            textBox2.Text = GetMenu.DFirst();
            textBox3.Text = GetMenu.DLast();
            textBox7.Text = GetMenu.DFirst();
            textBox8.Text = GetMenu.DLast();
            textBox10.Text = GetMenu.DFirst();
            textBox11.Text = GetMenu.DLast();
            sATT1DataGridView.Columns["NTD2"].HeaderText = "RMB";

     
        }


        private void GD4(string p)
        {
            throw new Exception("The method or operation is not implemented.");
        }
        public static System.Data.DataTable GetTT(string TTCode)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "SELECT ID,SUM(USDAMOUNT) USDAMOUNT,SUM(NTDAMOUNT) NTDAMOUNT FROM SATT2 where TTCode=@TTCode GROUP BY ID";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TTCode", TTCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }


        public static System.Data.DataTable GetTT2(string ID, string TTCode)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "SELECT ID,currency FROM SATT1 where seqno=@ID and TTCode=@TTCode";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
            command.Parameters.Add(new SqlParameter("@TTCode", TTCode));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "plc1");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["plc1"];
        }


        private void button2_Click(object sender, EventArgs e)
        {
            string TYPE = "";
            System.Data.DataTable dt1 = null;

            if (sATT1DataGridView.SelectedRows.Count == 0 || (textBox1.Text.ToString() == ""&& textBox9.Text.ToString() == ""))
            {
                MessageBox.Show("??????");
                return;
            }
            string da = sATT1DataGridView.SelectedRows[0].Cells["Seqno"].Value.ToString();

            if (globals.DBNAME == "CHOICE")
            {
                dt1 = GetMenu.GetttCHO(textBox1.Text);

            }
            string DOCENTRY = "";

 
                if (textBox1.Text != "")
                {
                    TYPE = "ORDR";
                    DOCENTRY = textBox1.Text.Trim();
                }
                else
                {
                    TYPE = "OINV";
                    DOCENTRY = textBox9.Text.Trim();
                }

                dt1 = GetMenu.Gettt(DOCENTRY, TYPE);
                if (dt1.Rows.Count == 0)
                {
                    TYPE = "OINV3";
                    dt1 = GetMenu.Gettt(DOCENTRY, TYPE);
                }

            
            System.Data.DataTable dt2 = sa.SATT2;

            for (int i = 0; i <= dt1.Rows.Count - 1; i++)
            {
                DataRow drw = dt1.Rows[i];
                DataRow drw2 = dt2.NewRow();
                

                drw2["ttcode"] = tTCodeTextBox.Text;
                drw2["id"] = da;
                string DOC = drw["docentry"].ToString();
                drw2["docentry"] = DOC;
                if (drw["????"].ToString() == "S")
                {
                    drw2["itemcode"] = drw["?y?z"];
                }
                else
                {
                    drw2["itemcode"] = drw["itemcode"];
                }
                string CURRENCY = drw["CURRENCY"].ToString();
                drw2["memo"] = drw["oinv"];
                drw2["quantity"] = drw["quantity"];
                drw2["price"] = drw["price"];
                drw2["usdamount"] = drw["PRICEAFVAT"];
                drw2["USDAMT1"] = drw["gtotalfc"];
                drw2["ntdamount"] = drw["gtotalC"];
                drw2["NTDAMT"] = drw["gtotalC"];
                    
       

                drw2["shipdate"] = drw["shipdate"];
                drw2["ttrate"] = drw["rate"];
                drw2["Tax"] = drw["vatprcnt"];
                drw2["USDTAX1"] = drw["rate"];
                drw2["cardcode"] = drw["cardcode"];
                drw2["cardname"] = drw["cardname"];
                drw2["LINENUM"] = drw["LineNum"];
                drw2["CardCode2"] = drw["?????s??"];
                drw2["CARDFNAME"] = drw["?^???W??"];

                System.Data.DataTable G3 = GetMenu.GetSA(DOC);
                if (G3.Rows.Count > 0)
                {
                    drw2["SA"] = G3.Rows[0]["?~??"].ToString();
                    drw2["SALES"] = G3.Rows[0]["?~??"].ToString();
                }
                dt2.Rows.Add(drw2);

            }
            textBox1.Text = "";
        }




        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuList();

            if (LookupValues != null)
            {


                System.Data.DataTable dt2 = sa.SATT1;

                DataRow drw2 = dt2.NewRow();
                drw2["cardname"] = Convert.ToString(LookupValues[1]);
                drw2["TTCode"] = tTCodeTextBox.Text;

                drw2["Seqno"] = sATT1DataGridView.Rows.Count.ToString();
                dt2.Rows.Add(drw2);


            }

        }



        private void sATT2DataGridView_DefaultValuesNeeded_1(object sender, DataGridViewRowEventArgs e)
        {
            if (sATT1DataGridView.SelectedRows.Count > 0)
            {
                string da = sATT1DataGridView.SelectedRows[0].Cells["Seqno"].Value.ToString();

                e.Row.Cells["dataGridViewTextBoxColumn3"].Value = da;
            }
            else
            {
                MessageBox.Show("????????");
            }
        }

        private void sATT2DataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {

            if (e.RowIndex >= sATT2DataGridView.Rows.Count)
                return;
            try
            {

                if (sATT1DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow dgr = sATT2DataGridView.Rows[e.RowIndex];
                    string da = sATT1DataGridView.SelectedRows[0].Cells["Seqno"].Value.ToString();
                    string dd = dgr.Cells["dataGridViewTextBoxColumn3"].Value.ToString();
                    if (da == dd)
                    {
                        dgr.DefaultCellStyle.BackColor = Color.Pink;
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void sATT1DataGridView_MouseClick_1(object sender, MouseEventArgs e)
        {
            sATT2TableAdapter.Fill(sa.SATT2, MyID);

            try
            {

                string da1 = sATT1DataGridView.SelectedRows[0].Cells["Seqno"].Value.ToString();
                for (int i = 0; i <= sATT2DataGridView.Rows.Count - 1; i++)
                {

                    DataGridViewRow row;

                    row = sATT2DataGridView.Rows[i];
                    string a0 = row.Cells["dataGridViewTextBoxColumn3"].Value.ToString();

                    if (da1 == a0)
                    {
                        sATT2DataGridView.FirstDisplayedScrollingRowIndex = i;
                        break;
                    }

                }
            }
            catch
            {

            }
        }

        private void sATT1DataGridView_DefaultValuesNeeded_1(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["SeqNo"].Value = util.GetSeqNo(2, sATT1DataGridView);
            e.Row.Cells["Company"].Value = "ACME";
            e.Row.Cells["Bank"].Value = "???n";
            e.Row.Cells["PAYCHECK"].Value = "TT";
            e.Row.Cells["Currency"].Value = "NTD";

        }

        private void sATT2DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (sATT2DataGridView.Columns[e.ColumnIndex].Name == "Quantity" ||
                          sATT2DataGridView.Columns[e.ColumnIndex].Name == "PriceTax" ||
                          sATT2DataGridView.Columns[e.ColumnIndex].Name == "TTRate" ||
                                     sATT2DataGridView.Columns[e.ColumnIndex].Name == "Price" ||
                          sATT2DataGridView.Columns[e.ColumnIndex].Name == "Tax")
                {

                    decimal iQuantity = 0;
                    decimal iUnitPrice = 0;
                    decimal iRate = 0;
                    decimal iTax = 0;
                    decimal iTax2 = 0;
                    iQuantity = Convert.ToInt32(this.sATT2DataGridView.Rows[e.RowIndex].Cells["Quantity"].Value);
                    iUnitPrice = Convert.ToDecimal(this.sATT2DataGridView.Rows[e.RowIndex].Cells["Price"].Value);
                    iRate = Convert.ToDecimal(this.sATT2DataGridView.Rows[e.RowIndex].Cells["TTRate"].Value);
                    iTax = Convert.ToDecimal(this.sATT2DataGridView.Rows[e.RowIndex].Cells["Tax"].Value);
                    iTax2 = iTax / 100 + 1;

                    this.sATT2DataGridView.Rows[e.RowIndex].Cells["NTDAmount"].Value = (iQuantity * iUnitPrice * iRate * iTax2).ToString("0");
                    this.sATT2DataGridView.Rows[e.RowIndex].Cells["NTDAMT"].Value = (iQuantity * iUnitPrice * iRate * iTax2).ToString("0");
                    this.sATT2DataGridView.Rows[e.RowIndex].Cells["USDAmount"].Value = (iQuantity * iUnitPrice * iTax2).ToString();
                    this.sATT2DataGridView.Rows[e.RowIndex].Cells["USDAMT"].Value = (iQuantity * iUnitPrice * iTax2).ToString();
                }


                if (sATT2DataGridView.Columns[e.ColumnIndex].Name == "USDAmount" ||
                     sATT2DataGridView.Columns[e.ColumnIndex].Name == "USDTAX")
                {

                    decimal USDAmount = 0;
                    decimal USDTAX = 0;

                    USDAmount = Convert.ToDecimal(this.sATT2DataGridView.Rows[e.RowIndex].Cells["USDAMT"].Value);
                    USDTAX = Convert.ToDecimal(this.sATT2DataGridView.Rows[e.RowIndex].Cells["USDTAX"].Value);

                    this.sATT2DataGridView.Rows[e.RowIndex].Cells["NTDAmount"].Value = (USDAmount * USDTAX).ToString("0");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }



        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                sATTTableAdapter.Fill(sa.SATT, MyID);
                sATT1TableAdapter.Fill(sa.SATT1, MyID);
                sATT2TableAdapter.Fill(sa.SATT2, MyID);

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            CalcTotals2();
        }

        private void CalcTotals2()
        {


            decimal NTD = 0;
            decimal USD = 0;


            int i = this.sATT2DataGridView.SelectedRows.Count - 1;
            for (int iRecs = 0; iRecs <= i; iRecs++)
            {
                if (!String.IsNullOrEmpty(sATT2DataGridView.SelectedRows[iRecs].Cells["USDAmount"].Value.ToString()))
                {
                    USD += Convert.ToDecimal(sATT2DataGridView.SelectedRows[iRecs].Cells["USDAmount"].Value);
                }
                else
                {
                    USD = 0;
                }

                if (!String.IsNullOrEmpty(sATT2DataGridView.SelectedRows[iRecs].Cells["NTDAmount"].Value.ToString()))
                {
                    NTD += Convert.ToDecimal(sATT2DataGridView.SelectedRows[iRecs].Cells["NTDAmount"].Value);
                }
                else
                {
                    NTD = 0;
                }

            }

            label1.Text = "?????`??: " + USD.ToString("#,##0.000");
            label2.Text = "?x???`??: " + NTD.ToString("#,##0.00");





        }

        private void button5_Click(object sender, EventArgs e)
        {

            try
            {
                if (sATT1DataGridView.SelectedRows.Count == 0)
                {
                    MessageBox.Show("??????????");
                    return;
                }
                string server = "//acmesrv01//SAP_Share//TTAdvance//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.download(filename);


                if (result == DialogResult.OK)
                {
                    MessageBox.Show(Path.GetFileName(opdf.FileName));
                    string file = opdf.FileName;
                    bool FF1 = getrma.UploadFile(file, server, false);
                    if (FF1 == false)
                    {
                        return;
                    }


                    DataGridViewRow row;

                    row = sATT1DataGridView.SelectedRows[0];
                    string a0 = row.Cells["Column1"].Value.ToString();
                    string a1 = row.Cells["ID"].Value.ToString();
                    string a2 = filename;

                    string a3 = @"\\acmesrv01\SAP_Share\TTAdvance\" + filename;


                    Updatepath(a2, a3, a0, a1);


                    sATT1TableAdapter.Fill(sa.SATT1, MyID);
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void sATT1DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "check2")
                {
                    for (int j = 0; j <= 1; j++)
                    {


                        System.Data.DataTable dt1 = sa.SATT1;
                        int i = e.RowIndex;
                        DataRow drw = dt1.Rows[i];

                        string aa = drw["path"].ToString();


                        System.Diagnostics.Process.Start(aa);


                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];

                        cell.LinkVisited = true;
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable GETACCCODE(string BANK, string CURRENCY, string PAY)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ACCCODE,ACCNAME FROM SATT_ACC WHERE BANK=@BANK AND CURRENCY =@CURRENCY AND PAY like '%" + PAY + "%' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BANK", BANK));
            command.Parameters.Add(new SqlParameter("@CURRENCY", CURRENCY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GETACCCODE2(string BANK)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT ACCCODE,ACCNAME FROM SATT_ACC WHERE (ACCNAME LIKE '%???s%' OR ACCNAME ='?O?s?g?a????#10196') AND BANK=@BANK ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BANK", BANK));
    
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GETCARDCODE(string TTCODE,string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(DOCENTRY)?@DOCENTRY,MAX(CARDNAME)?@CARDNAME,MAX(CARDCODE2) CARDCODE,SUM(USDAMOUNT)?@USD,AVG(TTRATE)?@RATE FROM SATT2 WHERE TTCODE=@TTCODE AND ID=@ID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TTCODE", TTCODE));
            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETSAPUSER()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT lastName+FIRSTNAME USERS FROM OHEM WHERE HOMETEL=@USERS");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GETMAXFINN()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT MAX(FinncPriod) FINN  FROM OBTF");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private void button6_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcelSelect(sATT2DataGridView);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView5.DataSource = Gettt1();

            System.Data.DataTable dtCost = MakeTableCombine();
            System.Data.DataTable dt = Gettt2();
            DataRow dr = null;
            string DuplicateKey = "";

            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                DataRow dd = dt.Rows[i];
                dr = dtCost.NewRow();
                dr["?I??????"] = dd["?I??????"].ToString();
                dr["?????W??"] = dd["?????W??"].ToString();
                dr["?P??????"] = dd["?P??????"].ToString();
                dr["AR????"] = dd["AR????"].ToString();
                dr["????"] = dd["????"].ToString();
                dr["???{????"] = dd["???{????"].ToString();
                dr["???q"] = dd["???q"].ToString();
                dr["????????"] = Convert.ToDecimal(dd["????????"]);
                dr["?????`?B"] = Convert.ToDecimal(dd["?????`?B"]);
                dr["???v"] = dd["???v"].ToString();
                dr["?x???`?B"] = dd["?q???`?B"].ToString();

                string CODE = dd["CODE"].ToString();

                if (CODE != DuplicateKey)
                {
                    string DD = dd["?J?b???B"].ToString();
                    dr["?q???t?B"] = Convert.ToDecimal(dd["?q???t?B"]);
                    dr["?J?b???B"] = Convert.ToDecimal(dd["?J?b???B"]);
                    dr["?????O"] = Convert.ToDecimal(dd["?????O"]);

                }

                DuplicateKey = CODE;
                dr["PO???X"] = dd["PO???X"].ToString();
                dr["?L?b????(AR)"] = dd["?L?b????"].ToString();
                dr["?I??????"] = dd["?I??????"].ToString();
                dr["?t????????"] = dd["?t????????"].ToString();
                dr["?????????O??????"] = dd["?????????O??????"].ToString();

                dtCost.Rows.Add(dr);

            }

            dataGridView1.DataSource = dtCost;
        }
        private System.Data.DataTable Gettt1()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct T2.TTDATE,Seqno,case isnull(t1.CardName,'') when '' then t0.cardname else t1.cardname end CardName,Company");
            sb.Append(" ,Bank,PAYCHECK,t0.Currency,Amount,CurrencyRate,Fee,");
            sb.Append(" TotalAmount,case when TTTotal2=0 then TTTotal when isnull(TTTotal2,0)=0 then TTTotal else TTTotal2 end TTTotal,Detail,t0.ttcode,t0.REMARK ????");
            sb.Append(" from satt1 t0 ");
            sb.Append(" left join satt2 t1 on(t0.ttcode=t1.ttcode and t0.seqno=t1.id)");
            sb.Append(" left join satt t2 on(t0.ttcode=t2.ttcode)  ");
            sb.Append(" left join acmesql02.dbo.dln1 t4 on (t4.baseentry=T1.docentry and  t4.baseline=t1.linenum )");
            sb.Append(" left join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'  )   ");
     
            sb.Append("  where  1=1 ");
            if (textBox4.Text != "")
            {
                sb.Append(" and  t1.[cardname] like N'%" + textBox4.Text.ToString() + "%'  ");
            }
            if (textBox2.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and T2.TTDATE between @DocDate1 and @DocDate2");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and ISNULL(DETAIL,'') = '' ");

            }
            if (checkBox2.Checked)
            {
                sb.Append(" and ISNULL(tttotal,1) <> 0 ");

            }
            if (checkBox4.Checked)
            {
                sb.Append("  AND substring(t1.cardcode2,1,1)=8 ");

            }
            sb.Append(" order by T2.TTDATE  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox2.Text.ToString()));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox3.Text.ToString()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable GETARPAY()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select (T1.TTDATE) ????????,MAX(T6.CARDCODE) ?????s??,MAX(T0.CardName) ?????W??,MAX(T2.Bank) ?q??????,MAX(T2.PAYCHECK) ?I??????,");
            sb.Append(" MAX(T2.Currency)  ???O,T7.DOCENTRY ?P???q??,T6.DOCENTRY AR????,SUM(t0.Quantity) ???q,SUM(USDAMT) '?????J?b???B(????)',MAX(isnull(USDTAX,0)) ???v,");
            sb.Append(" SUM(NTDAmount) '?????J?b???B(?x??)',MAX(Convert(varchar(8),T6.DOCDATE,112)) '?L?b????(AR)',MAX(T6.U_ACME_PAY) ?I??????,");
            sb.Append(" MAX(Convert(varchar(8), ACMESQL02.dbo.fun_CreditDate(T6.u_acme_pay,T0.CardCode2,T6.DocDate),112))  ?t????????");
            sb.Append(" ,MAX(DATEDIFF(D,ACMESQL02.dbo.fun_CreditDate(T6.u_acme_pay,T0.CardCode2,T6.DocDate),CAST(T1.TTDATE AS DATETIME))) ?????????O??????           ");
            sb.Append(" ,MAX(T0.SA) ?~??,MAX(T0.SALES) ?~?? from SATT T1    ");
            sb.Append(" LEFT JOIN SATT1 T2 ON (T1.TTCODE=T2.TTCODE)  ");
            sb.Append(" LEFT JOIN satt2 t0  ON (T0.TTCODE=T2.TTCODE AND T0.ID=T2.Seqno)  ");
            sb.Append(" left join acmesql02.dbo.dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum and t4.basetype='17'  )   ");
            sb.Append(" left join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'  )   ");
            sb.Append(" left join acmesql02.dbo.OINV T6 on (t5.DOCENTRY=T6.DOCENTRY)          ");
            sb.Append(" left join acmesql02.dbo.RDR1 T7 on (T0.docentry=T7.DOCENTRY AND T0.linenum=T7.linenum)      ");
            sb.Append(" WHERE Convert(varchar(8),T6.DOCDATE,112)  BETWEEN @DocDate1 and @DocDate2 ");
            sb.Append(" GROUP BY T7.DOCENTRY,T6.DOCENTRY,T1.TTDATE");
            sb.Append(" ORDER BY T1.TTDATE,T6.DOCENTRY");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox10.Text));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox11.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

  
        private System.Data.DataTable Gettteu(string COMPANY)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();



            sb.Append(" SET LANGUAGE N'Traditional Chinese' ");
            sb.Append(" select DISTINCT T2.CARDCODE2 ???X,T1.TTDATE ????????, DATENAME(weekday,CAST(T1.TTDATE AS DATETIME)) ?P??,");
            sb.Append(" CASE WHEN T2.CARDCODE2 IN ('0511-00','0257-00') THEN T0.CARDNAME ELSE  ");
            sb.Append(" CASE ISNULL(T2.CARDNAME,'') WHEN '' THEN T0.CARDNAME ELSE T2.CARDNAME END END ?????W?? ");
            sb.Append(" ,t0.ttcode code,SEQNO,Bank ???J????,CurrencyRate ???v, ");
            sb.Append(" isnull(case T0.Currency when 'USD' THEN CAST(Amount AS VARCHAR) ELSE '' END,0) USD, ");
            sb.Append(" isnull(case T0.Currency when 'NTD' THEN CAST(Amount AS VARCHAR) ELSE '' END,0) NTD,isnull(Fee,0) ?????O,PAYCHECK PT,CURRENCY CU,isnull(Amount,0) OC from satt1 t0 left join satt t1 on (t0.ttcode=t1.ttcode) ");
            sb.Append(" left join satt2 t2 on (t2.ttcode=t0.ttcode AND T2.ID=T0.SEQNO)");
            sb.Append(" WHERE T1.TTDATE between  @DocDate1 and @DocDate2 AND COMPANY=@COMPANY and isnull(alcheck,'')='' ORDER BY T1.TTDATE ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox2.Text.ToString()));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox3.Text.ToString()));
            command.Parameters.Add(new SqlParameter("@COMPANY", COMPANY));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable NA1(string TTCODE, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(DOCTOTAL) FROM  ACMESQL02.DBO.OINV WHERE ");
            sb.Append(" DOCENTRY IN (");
            sb.Append("         select DISTINCT case isnull(t0.memo,'') when '' then t5.docentry else t0.memo end AR???? from satt2 t0");
            sb.Append("  left join acmesql02.dbo.dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum)");
            sb.Append("   left join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'  )  WHERE TTCODE=@TTCODE AND ID=@ID)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TTCODE", TTCODE));
            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable NAH(string ID1)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("         select DISTINCT t5.docentry  AR???? from satt2 t0");
            sb.Append("  left join acmesql02.dbo.dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum)");
            sb.Append("   left join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'  )  WHERE ID1=@ID1 ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID1", ID1));


            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable NA2(string TTCODE, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
          //  sb.Append(" SELECT SUM(USDAMOUNT) FROM SATT2 WHERE TTCODE=@TTCODE AND ID=@ID ");
            sb.Append(" SELECT SUM(USDAMOUNT)   FROM SATT2 WHERE MEMO IN (SELECT MEMO  FROM SATT2 WHERE  TTCODE=@TTCODE AND ID=@ID ) ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TTCODE", TTCODE));
            command.Parameters.Add(new SqlParameter("@ID", ID));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
    
        private System.Data.DataTable Gettteu2()
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SET LANGUAGE N'Traditional Chinese' ");
            sb.Append(" select DISTINCT T2.CARDCODE2 ???X,T1.TTDATE ????????,DATENAME(weekday,CAST(T1.TTDATE AS DATETIME)) ?P??,t0.ttcode code,SEQNO,Bank ???J????,CurrencyRate ???v,T0.CARDNAME ?????W??2, ");
            sb.Append(" CASE WHEN T2.CARDCODE2 IN ('0511-00','0257-00')  AND COMPANY='ACME' THEN T0.CARDNAME ELSE  ");
            sb.Append(" CASE ISNULL(T2.CARDNAME,'') WHEN '' THEN T0.CARDNAME ELSE T2.CARDNAME END END ?????W??, ");
            sb.Append(" case T0.Currency when 'USD' THEN CAST(CurrencyRate*TotalAmount AS INT) ELSE TotalAmount END NTD, ");
            sb.Append(" case T0.Currency when 'NTD' THEN CAST(ROUND(TotalAmount/case CurrencyRate when 0 then TotalAmount else CurrencyRate end,3) AS DECIMAL(12,3)) ELSE TotalAmount END USD,cast(amount*CurrencyRate as decimal(12,3)) ?J?b???B,cast(Fee*CurrencyRate as decimal(12,3)) ?????O,cast(CurrencyRate*TotalAmount as decimal(12,3)) ???????B      ");
            sb.Append(" ,TTCheck ?q???O,CASE BANKCHECK  WHEN 'TRUE' THEN '??????' END  ??????,T3.CARDFNAME ?~???W??,T5.lastName +T5.firstName ?~?U   from satt1 t0 left join satt t1 on (t0.ttcode=t1.ttcode) ");
            sb.Append(" left join satt2 t2 on (t2.ttcode=t0.ttcode AND T2.ID=T0.SEQNO)  ");
            sb.Append(" left join ACMESQL02.DBO.OCRD t3 on (t2.CARDCODE2=t3.CARDCODE COLLATE  Chinese_Taiwan_Stroke_CI_AS) ");
            sb.Append(" left join ACMESQL02.DBO.ORDR T4 on (T2.DOCENTRY=T4.DOCENTRY ) ");
            sb.Append(" left join ACMESQL02.DBO.OHEM T5 on (T4.OwnerCode =T5.EMPID ) ");
                if (checkBox3.Checked)
                {
                    sb.Append(" WHERE T1.TTDATE BETWEEN @DocDate2 AND @DocDate3 order by Seqno   ");
                }
                else
                {
                    sb.Append(" WHERE T1.TTDATE = @DocDate1 order by Seqno   ");
                }
            

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", tTDateTextBox.Text.ToString()));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox2.Text.ToString()));
            command.Parameters.Add(new SqlParameter("@DocDate3", textBox3.Text.ToString()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
   
        private System.Data.DataTable Gettteu1(string ttcode, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT DOCENTRY  ???? FROM SATT2  WHERE ttcode = @ttcode  AND ID=@ID ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ttcode", ttcode));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Gettteu2(string ttcode, string ID)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select DISTINCT case isnull(t0.memo,'') when '' then t5.docentry else t0.memo end AR???? from satt2 t0");
            sb.Append(" left join acmesql02.dbo.dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum)");
            sb.Append("  left join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'  )  WHERE T0.ttcode = @ttcode  AND T0.ID=@ID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ttcode", ttcode));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Gettt2()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select t2.ttcode,T2.SEQNO,T1.TTDATE ?I??????,t0.CardName ?????W??,t0.Docentry ?P??????,case  when  isnull(t0.memo,'') = '' then t5.docentry else t0.memo end AR????,");
            sb.Append(" CASE WHEN t0.ItemCode LIKE '%?`??%' THEN T7.Dscription  ELSE t0.ItemCode END COLLATE  Chinese_Taiwan_Stroke_CI_AS  ????");
            sb.Append(" ,t0.ShipDate ???{????,isnull(t0.Quantity,0) ???q,isnull(t0.Price,0) ????????,isnull(Tax,0) Tax,isnull(USDAmount,0) ?????`?B,isnull(TTRate,0) ???v,  ");
            sb.Append(" isnull(t0.NTDAmount,0) ?q???`?B,CASE ISNULL(TTTotal2,0) WHEN 0 THEN isnull(TTTotal,0) ELSE isnull(TTTotal2,0) END ?q???t?B,isnull(t2.Amount,0) ?J?b???B,isnull(t2.fee,0) ?????O,t0.CardCode PO???X,t2.ttcode +CAST(SEQNO AS NVARCHAR) CODE ");
            sb.Append(" ,Convert(varchar(8),T6.DOCDATE,112) ?L?b????,T6.U_ACME_PAY ?I??????, ");
            sb.Append(" Convert(varchar(8), ACMESQL02.dbo.fun_CreditDate(T6.u_acme_pay,T0.CardCode2,T6.DocDate),112)  ?t???????? ");
            sb.Append(" ,DATEDIFF(D,ACMESQL02.dbo.fun_CreditDate(T6.u_acme_pay,T0.CardCode2,T6.DocDate),CAST(T1.TTDATE AS DATETIME)) ?????????O?????? ");
            sb.Append(" from satt2 t0    ");
            sb.Append(" left join acmesql02.dbo.dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum )  ");
            sb.Append(" left join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'   )  ");
            sb.Append(" left join acmesql02.dbo.OINV T6 on (t5.DOCENTRY=T6.DOCENTRY)      ");
            sb.Append(" left join acmesql02.dbo.RDR1 T7 on (T0.docentry=T7.DOCENTRY AND T0.linenum=T7.linenum)           ");
            sb.Append(" left join satt t1 on (t0.ttcode=t1.ttcode)   ");
            sb.Append(" left join satt1 t2 on (t0.ttcode=t2.ttcode AND T2.SEQNO=T0.ID)      ");
            sb.Append(" left JOIN acmesql02.dbo.OITM T11 ON t4.ITEMCODE = T11.ITEMCODE ");
            sb.Append("   where  1=1 ");
            if (textBox4.Text != "")
            {
                sb.Append(" and  t0.[cardname] like N'%" + textBox4.Text.ToString() + "%'  ");
            }
            if (textBox2.Text != "" && textBox3.Text != "")
            {
                sb.Append(" and T1.TTDATE between @DocDate1 and @DocDate2");
            }
            if (checkBox1.Checked)
            {
                sb.Append(" and ISNULL(DETAIL,'') = '' ");

            }
            if (checkBox2.Checked)
            {
                sb.Append(" and ISNULL(tttotal,1) <> 0 ");

            }
            if (checkBox4.Checked)
            {
                sb.Append("  AND substring(T0.cardcode2,1,1)=8 ");

            }

            sb.Append(" ORDER BY TTCODE,SEQNO");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox2.Text.ToString()));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox3.Text.ToString()));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private System.Data.DataTable Gettt3(string ttcode, string id)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select ID,T0.CardName,t0.Docentry,T0.linenum,");
            sb.Append(" CASE WHEN t0.ItemCode LIKE '%?`??%' THEN T7.Dscription  ELSE t0.ItemCode END COLLATE  Chinese_Taiwan_Stroke_CI_AS ItemCode ,");
            sb.Append(" t0.ShipDate,t0.Quantity,t0.Price,Tax,USDAmount,TTRate,  ");
            sb.Append(" t0.NTDAmount,t0.CardCode,case  when  isnull(t0.memo,'') = '' then t5.docentry else t0.memo end memo ");
            sb.Append(" ,Convert(varchar(8),T6.DOCDATE,112) ?L?b????,T6.U_ACME_PAY ?I??????, ");
            sb.Append(" Convert(varchar(8), ACMESQL02.dbo.fun_CreditDate(T6.u_acme_pay,T0.CardCode2,T6.DocDate),112)  ?t???????? ");
            sb.Append(" ,DATEDIFF(D,ACMESQL02.dbo.fun_CreditDate(T6.u_acme_pay,T0.CardCode2,T6.DocDate),CAST(T1.TTDATE AS DATETIME)) ?????????O?????? from satt2 t0    ");
            sb.Append(" LEFT JOIN SATT T1 ON (T0.TTCODE=T1.TTCODE) ");
            sb.Append(" left join acmesql02.dbo.dln1 t4 on (t4.baseentry=T0.docentry and  t4.baseline=t0.linenum and t4.basetype='17'  )  ");
            sb.Append(" left join acmesql02.dbo.inv1 t5 on (t5.baseentry=T4.docentry and  t5.baseline=t4.linenum and t5.basetype='15'  )  ");
            sb.Append(" left join acmesql02.dbo.OINV T6 on (t5.DOCENTRY=T6.DOCENTRY)         ");
            sb.Append(" left join acmesql02.dbo.RDR1 T7 on (T0.docentry=T7.DOCENTRY AND T0.linenum=T7.linenum)     ");
            sb.Append(" WHERE  T0.ttcode=@ttcode and id=@id ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);


            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ttcode", ttcode));
            command.Parameters.Add(new SqlParameter("@id", id));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }


        private void button7_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void dataGridView5_MouseClick(object sender, MouseEventArgs e)
        {
            if (dataGridView5.SelectedRows.Count > 0)
            {
                string da = dataGridView5.SelectedRows[0].Cells["Seqno1"].Value.ToString();
                string da1 = dataGridView5.SelectedRows[0].Cells["ttcode"].Value.ToString();
                dataGridView6.DataSource = Gettt3(da1, da);
            }
        }
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("?q???t?B", typeof(Decimal));
            dt.Columns.Add("?I??????", typeof(string));
            dt.Columns.Add("?????W??", typeof(string));
            dt.Columns.Add("?P??????", typeof(string));
            dt.Columns.Add("AR????", typeof(string));
            dt.Columns.Add("????", typeof(string));
            dt.Columns.Add("???{????", typeof(string));
            dt.Columns.Add("???q", typeof(string));
            dt.Columns.Add("????????", typeof(Decimal));
            dt.Columns.Add("?????`?B", typeof(Decimal));
            dt.Columns.Add("???v", typeof(string));
            dt.Columns.Add("?x???`?B", typeof(string));
            dt.Columns.Add("?J?b???B", typeof(Decimal));
            dt.Columns.Add("?????O", typeof(Decimal));
            dt.Columns.Add("PO???X", typeof(string));
            dt.Columns.Add("?L?b????(AR)", typeof(string));
            dt.Columns.Add("?I??????", typeof(string));
            dt.Columns.Add("?t????????", typeof(string));
            dt.Columns.Add("?????????O??????", typeof(string));

            return dt;
        }
        private System.Data.DataTable MakeTableCombineEu()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("???X", typeof(string));
            dt.Columns.Add("?????W??", typeof(string));
            dt.Columns.Add("????????", typeof(string));
            dt.Columns.Add("?P??", typeof(string));
            dt.Columns.Add("?P??????", typeof(string));
            dt.Columns.Add("AR????", typeof(string));
            dt.Columns.Add("???J????", typeof(string));
            dt.Columns.Add("???v", typeof(string));
            dt.Columns.Add("NTD", typeof(string));
            dt.Columns.Add("USD", typeof(string));
            dt.Columns.Add("?J?b???B", typeof(string));
            dt.Columns.Add("?????O", typeof(string));
            dt.Columns.Add("PT", typeof(string));
            dt.Columns.Add("CU", typeof(string));
            dt.Columns.Add("OC", typeof(string));
            dt.Columns.Add("ARUSD", typeof(string));
            dt.Columns.Add("ARNTD", typeof(string));
            dt.Columns.Add("AREX", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombineEu2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("???X", typeof(string));
            dt.Columns.Add("?????W??", typeof(string));
            dt.Columns.Add("?????W??2", typeof(string));
            dt.Columns.Add("????????", typeof(string));
            dt.Columns.Add("?P??", typeof(string));
            dt.Columns.Add("???J????", typeof(string));
            dt.Columns.Add("NTD", typeof(string));
            dt.Columns.Add("???v", typeof(string));
            dt.Columns.Add("USD", typeof(string));
            dt.Columns.Add("?P??????", typeof(string));
            dt.Columns.Add("AR????", typeof(string));
            dt.Columns.Add("????", typeof(string));
            dt.Columns.Add("?J?b???B", typeof(string));
            dt.Columns.Add("?????O", typeof(string));
            dt.Columns.Add("???????B", typeof(string));
            dt.Columns.Add("?q???O", typeof(string));
            dt.Columns.Add("??????", typeof(string));
            dt.Columns.Add("?~???W??", typeof(string));
            dt.Columns.Add("?~?U", typeof(string));
            
            return dt;
        }
        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = GetLC1();
        }

        private System.Data.DataTable GetLC1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT lccode,iSSUEDATE ????,LCNO LC,CARDNAME ????,COMPANY ???q,");
            sb.Append(" LCAMT LC???B,LCTOTAL ?R?P???B,LCFINAL ?l?B,EXPIRY,shipdate FROM account_LC where 1=1 ");

            if (textBox5.Text != "")
            {
                sb.Append(" and  CARDNAME like '%" + textBox5.Text.ToString() + "%'  ");
            }
            if (textBox7.Text != "" && textBox8.Text != "")
            {
                sb.Append(" and iSSUEDATE between @DocDate1 and @DocDate2");
            }
            if (textBox6.Text != "")
            {
                sb.Append(" and LCNO  like '%" + textBox6.Text.ToString() + "%' ");
            }
            sb.Append(" order by iSSUEDATE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocDate1", textBox7.Text.ToString()));
            command.Parameters.Add(new SqlParameter("@DocDate2", textBox8.Text.ToString()));


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
        private System.Data.DataTable GetLC2(string lccode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select seqno ????,Model ?~?W,Quantity ???q,Price ????,amount ???B,Quantity1 ???q1,amount1 ???B1 from account_LC1 where lccode=@lccode order by  seqno");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@lccode", lccode));

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
        private System.Data.DataTable GetLC3(string lccode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("      select seqno ????,Model ?~?W,Quantity ???q,Price ????,amount ???B,BANK ?????? ");
            sb.Append(",LCTTPE ???O,LCDATE ??????,INDATE ?w?p?e?I??,REDATE ?????J?b??,MEMO LC?????O??  from account_LC2 where lccode=@lccode order by  seqno");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@lccode", lccode));

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

        private void dataGridView2_MouseClick(object sender, MouseEventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                daXX = dataGridView2.SelectedRows[0].Cells["lccode"].Value.ToString();

                dataGridView3.DataSource = GetLC2(daXX);
                dataGridView4.DataSource = GetLC3(daXX);

                LCXX = GetMenu.Account_LCDownload(daXX, "1");
                if (LCXX.Rows.Count > 0)
                {
                    button10.Visible = true;

                }
                else
                {
                    button10.Visible = false;
                }

                LCXX2 = GetMenu.Account_LCDownload(daXX, "2");
                if (LCXX2.Rows.Count > 0)
                {
                    button21.Visible = true;

                }
                else
                {
                    button21.Visible = false;
                }

            }
        }



        private void button10_Click(object sender, EventArgs e)
        {


            DataRow drw = LCXX.Rows[0];
            string aa = drw["filepath"].ToString();
            System.Diagnostics.Process.Start(aa);

        }


        private void UpdateSQL(string docentry, string linenum, string ardocentry)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" update  safor set ardocentry=@ardocentry where docentry=@aa and linenum=@bb ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@aa", docentry));
            command.Parameters.Add(new SqlParameter("@bb", linenum));
            command.Parameters.Add(new SqlParameter("@ardocentry", ardocentry));
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


        private void sATT1DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (sATT1DataGridView.Columns[e.ColumnIndex].Name == "Amount" ||
                          sATT1DataGridView.Columns[e.ColumnIndex].Name == "CurrencyRate")
                {


                    string FF = this.sATT1DataGridView.Rows[e.RowIndex].Cells["CurrencyRate"].Value.ToString();
                    if (String.IsNullOrEmpty(FF))
                    {
                        FF = "1";
                    }
                    decimal Amount = 0;
                    decimal CurrencyRate = 0;
                    Amount = Convert.ToDecimal(this.sATT1DataGridView.Rows[e.RowIndex].Cells["Amount"].Value);
                    CurrencyRate = Convert.ToDecimal(FF);


                    this.sATT1DataGridView.Rows[e.RowIndex].Cells["NTD2"].Value = (Amount * CurrencyRate).ToString("0");

                }
                if (sATT1DataGridView.Columns[e.ColumnIndex].Name == "Amount" ||
                          sATT1DataGridView.Columns[e.ColumnIndex].Name == "Fee")
                {

                    string FF = this.sATT1DataGridView.Rows[e.RowIndex].Cells["Fee"].Value.ToString();
                    if (String.IsNullOrEmpty(FF))
                    {
                        FF = "0";
                    }
                    decimal Amount = 0;
                    decimal Fee = 0;
                    Amount = Convert.ToDecimal(this.sATT1DataGridView.Rows[e.RowIndex].Cells["Amount"].Value);
                    Fee = Convert.ToDecimal(FF);

                    this.sATT1DataGridView.Rows[e.RowIndex].Cells["TotalAmount"].Value = (Amount + Fee).ToString();

                }


                if (sATT1DataGridView.Columns[e.ColumnIndex].Name == "SAPCHECK")
                {
                    string FF = this.sATT1DataGridView.Rows[e.RowIndex].Cells["SAPCHECK"].Value.ToString();

                    if (FF == "1")
                    {
                        DialogResult result;
                        result = MessageBox.Show("???T?{?O?_?n????SAP", "YES/NO", MessageBoxButtons.YesNo);
                        if (result == DialogResult.Yes)
                        {
                            System.Data.DataTable S1 = GETSAPUSER();
                            if (S1.Rows.Count > 0)
                            {
                                int n;
                                decimal n2;
                                string USER = S1.Rows[0][0].ToString();
                                string Seqno = this.sATT1DataGridView.Rows[e.RowIndex].Cells["Seqno"].Value.ToString();
                                string Amount = this.sATT1DataGridView.Rows[e.RowIndex].Cells["Amount"].Value.ToString();
                                 string CR = this.sATT1DataGridView.Rows[e.RowIndex].Cells["CurrencyRate"].Value.ToString();
                                string MEMO = "";
                                string MEMO2 = "";
                                string MEMO3 = "";
                                decimal AMTUSD = 0;
                                decimal AMTRATE = 0;
                                System.Data.DataTable L1 = GETCARDCODE(tTCodeTextBox.Text, Seqno);
                                string DOCENTRY = L1.Rows[0]["DOCENTRY"].ToString();
                              
                                    string CARDCODE = L1.Rows[0]["CARDCODE"].ToString();
                                    string CARDNAME = L1.Rows[0]["CARDNAME"].ToString();
                                    if (!String.IsNullOrEmpty(L1.Rows[0]["USD"].ToString()))
                                    {
                                        AMTUSD = Convert.ToDecimal(L1.Rows[0]["USD"]);
                                        AMTRATE = Convert.ToDecimal(L1.Rows[0]["RATE"]);
                                    }
                                    if (String.IsNullOrEmpty(CARDNAME))
                                    {
                                        //CRDNAME
                                        CARDNAME = sATT1DataGridView.Rows[e.RowIndex].Cells["CRDNAME"].Value.ToString();
                                    }
                                    CARDNAME = util.CARDNAME(CARDNAME);

                                    string DDDTIIME = Convert.ToInt16(tTDateTextBox.Text.Substring(4, 2)).ToString() + "/" + Convert.ToInt16(tTDateTextBox.Text.Substring(6, 2)).ToString();
                                    string DOC = DOCENTRY + "/";
                                    if (String.IsNullOrEmpty(DOCENTRY))
                                    {
                                        DOC = "";
                                    }
                                    MEMO = DOC + CARDCODE + CARDNAME + DDDTIIME + "?s";
                                    MEMO2 = MEMO;
                                    MEMO3 = MEMO;
                                
                                string BANK = this.sATT1DataGridView.Rows[e.RowIndex].Cells["Bank"].Value.ToString();
                                string Currency = this.sATT1DataGridView.Rows[e.RowIndex].Cells["Currency"].Value.ToString();
                                string PAY = this.sATT1DataGridView.Rows[e.RowIndex].Cells["PAYCHECK"].Value.ToString();
                                decimal NTD = Convert.ToDecimal(this.sATT1DataGridView.Rows[e.RowIndex].Cells["NTD2"].Value);
                                decimal CurrencyRate = 1;
                                string FFEE = this.sATT1DataGridView.Rows[e.RowIndex].Cells["Fee"].Value.ToString();
                                decimal Fee = 0;
                                if (!String.IsNullOrEmpty(CR))
                                {
                                    CurrencyRate = Convert.ToDecimal(CR);
                                }
                                if (decimal.TryParse(FFEE, out n2))
                                {
                                  
                                    Fee = Convert.ToDecimal(FFEE) * CurrencyRate;

                                }
                                decimal F3 = Math.Round(Fee, 0, MidpointRounding.AwayFromZero);
                                decimal F4 = NTD + F3;
                                if (Currency != "NTD")
                                {

                                    string CCR = CurrencyRate.ToString("G29");
                                    string AMT = Convert.ToDecimal(Amount).ToString("G29");
                                    if (F3 != 0)
                                    {

                                        MEMO3 = MEMO + "US" + (Convert.ToDecimal(Amount) + Convert.ToDecimal(FFEE)).ToString("G29") + "*" + CCR;

                                        MEMO2 = MEMO + "US" + Convert.ToDecimal(FFEE).ToString("G29") + "*" + CCR + "??????";
                                        MEMO = MEMO + "US" + AMT + "*" + CCR;
                                    }
                                    else
                                    {
                                        MEMO = MEMO + "US" + AMT + "*" + CCR;
                                        MEMO3 = MEMO;
                                    }
                                }
                                else
                                {
                                    if (AMTUSD != 0 && AMTRATE != 0)
                                    {
                                        MEMO3 = MEMO + "=US" + AMTUSD.ToString("G29") + "*" + AMTRATE.ToString("G29");
                                    }
                                }


    
                                string BU = this.sATT1DataGridView.Rows[e.RowIndex].Cells["BU"].Value.ToString();
                                string NFEE = this.sATT1DataGridView.Rows[e.RowIndex].Cells["TTCheck"].Value.ToString();
                                decimal NN = 0;
                                
                                if (int.TryParse(NFEE, out n))
                                {
                                    if (NFEE != "0")
                                    {
                                        NN = Convert.ToDecimal(NFEE);
                                    }
                                }
                                decimal F5 = F4 + NN;
                                //Fee
                                System.Data.DataTable S2 = GETACCCODE(BANK, Currency, PAY);
                                if (S2.Rows.Count > 0)
                                {
                                    string ACCCODE = S2.Rows[0][0].ToString();
                                    int FINN = Convert.ToInt16(GETMAXFINN().Rows[0][0]);
                                     if (FF == "1")
                                    {
                                        //ID
                                        string ID = this.sATT1DataGridView.Rows[e.RowIndex].Cells["ID"].Value.ToString();
                                        util.AddOBTD(F5, 28);
                                        util.AddOBTF(USER, F5, FINN, 28, ID);

                                        util.AddBTF1(0, ACCCODE, NTD, 0, MEMO, "-1", USER, "", BU, 28, FINN, "", "N", "D");
                                        if (F3 != 0)
                                        {
                                            util.AddBTF1(1, "62280201", F3, 0, MEMO2, "-1", USER, "", BU, 28, FINN, "", "N", "D");
                                        }
                                        util.AddBTF1(2, "22610102", 0, F4, MEMO3, "-1", USER, "", BU, 28, FINN, "", "N", "C");
                                        System.Data.DataTable S3 = GETACCCODE2(BANK);
                                        if (S3.Rows.Count > 0)
                                        {
                                            //62280201
                                            if (int.TryParse(NFEE, out n))
                                            {
                                                if (NFEE != "0")
                                                {
                                                    string ACCCODE2 = S3.Rows[0][0].ToString();
                                                    MEMO = MEMO3 + "?q???O";
                                                    util.AddBTF1(3, "62280201", NN, 0, MEMO, "-1", USER, "", BU, 28, FINN, "", "N", "D");
                                                    util.AddBTF1(4, ACCCODE2, 0, NN, MEMO, "-1", USER, "", BU, 28, FINN, "", "N", "C");
                                                }

                                            }
                                        }
                                        string T1 = util.GetONNM2().Rows[0][0].ToString();
                                        util.ADDONNM();
             

                                        MessageBox.Show("???J???\ ???O?b???????X : " + T1);

                                    }
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void sATT2DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void sATT1DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void button20_Click(object sender, EventArgs e)
        {
   

            if (textBox2.Text == "" || textBox3.Text == "")
            {

                MessageBox.Show("?????J????????");
                return;
            }
            if (checkBox3.Checked)
            {
                Execu();
            }
            else
            {
                System.Data.DataTable DT1 = DT("ACME");

                //    DT1.DefaultView.Sort = "????????";
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\ACC\\TT.xls";
                string ExcelTemplate = FileName;

                string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                      DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                //???? Excel ReportdataGridView1
                ExcelReport.NANCY(DT1, ExcelTemplate, OutPutFile);
            }
            checkBox3.Checked = false;
        }


        System.Data.DataTable DT(string ss)
        {
            System.Data.DataTable dtCost = MakeTableCombineEu();
            System.Data.DataTable dt = Gettteu(ss);
            System.Data.DataTable dt1 = null;
            System.Data.DataTable dt2 = null;

            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                DataRow dd = dt.Rows[i];
                dr = dtCost.NewRow();

                string gg = dd["code"].ToString();
                string SEQNO = dd["SEQNO"].ToString();
                dr["???X"] = dd["???X"].ToString();
                dr["?????W??"] = dd["?????W??"].ToString();

                dr["???J????"] = dd["???J????"].ToString();
                dr["PT"] = dd["PT"].ToString();
                dr["CU"] = dd["CU"].ToString();
                dr["OC"] = dd["OC"].ToString();
                dr["NTD"] = dd["NTD"].ToString();
                dr["?????O"] = dd["?????O"].ToString();
                dr["????????"] = dd["????????"].ToString();
                dr["?P??"] = dd["?P??"].ToString();

                string RATE = dd["???v"].ToString();
                dr["???v"] = RATE;

                dr["USD"] = dd["USD"].ToString();
                dt1 = Gettteu1(gg, SEQNO);
                StringBuilder sb = new StringBuilder();
                for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                {
                    DataRow dv = dt1.Rows[j];
                    string GH = dv["????"].ToString();
                    if (!String.IsNullOrEmpty(GH))
                    {
                        sb.Append(GH + "&");

                    }
                }

                if (!String.IsNullOrEmpty(sb.ToString()))
                {
                    sb.Remove(sb.Length - 1, 1);
                }
                string df = "";
                string docentry = sb.ToString();
                if (!String.IsNullOrEmpty(docentry))
                {
                    df = docentry + "/";
                }
                dr["?P??????"] = docentry;
                string SD = "";

                if (!String.IsNullOrEmpty(dr["?????W??"].ToString()))
                {
                    SD = dr["?????W??"].ToString().Substring(0, 2);
                }


                dt2 = Gettteu2(gg, SEQNO);

                StringBuilder sJ = new StringBuilder();
                for (int K = 0; K <= dt2.Rows.Count - 1; K++)
                {
                    DataRow dK = dt2.Rows[K];
                    string fg = dK["AR????"].ToString();
                    if (!String.IsNullOrEmpty(fg))
                    {
                        sJ.Append(fg + "&");
                    }

                }

                if (!String.IsNullOrEmpty(sJ.ToString()))
                {
                    sJ.Remove(sJ.Length - 1, 1);
                }

                dr["AR????"] = sJ.ToString();


                System.Data.DataTable N1 = NA1(gg, SEQNO);
                if (N1.Rows.Count > 0)
                {
                    dr["ARNTD"] = N1.Rows[0][0].ToString();
                }
                //System.Data.DataTable N2 = NA2(gg, SEQNO);
                //if (N2.Rows.Count > 0)
                //{
                //    dr["ARUSD"] = N2.Rows[0][0].ToString();
                //}
                sd = 0;
                System.Data.DataTable dt1h = GetOrderDataAP1(gg, SEQNO);
                for (int j = 0; j <= dt1h.Rows.Count - 1; j++)
                {
                    DataRow dd1 = dt1h.Rows[j];

              

                    if ((!String.IsNullOrEmpty(dd1["???q"].ToString())) && (!String.IsNullOrEmpty(dd1["????????"].ToString())) && (!String.IsNullOrEmpty(dd1["?|?v"].ToString())))
                    {

                        sd += Convert.ToDecimal(dd1["???q"]) * Convert.ToDecimal(dd1["????????"]) * Convert.ToDecimal(dd1["?|?v"]);

                    }
                }
                string usd = sd.ToString();
                dr["ARUSD"] = usd;

                string A1 = dr["ARUSD"].ToString();
                string A2 = dr["ARNTD"].ToString();
                if (!String.IsNullOrEmpty(dr["ARUSD"].ToString()) && !String.IsNullOrEmpty(dr["ARNTD"].ToString()))
                {
                    try
                    {
                        dr["AREX"] = Convert.ToString(Convert.ToDecimal(dr["ARNTD"].ToString()) / Convert.ToDecimal(dr["ARUSD"].ToString()));
                    }
                    catch
                    {

                    }
                }
                dtCost.Rows.Add(dr);

            }





            return dtCost;
        }
        private void Execu()
        {

            System.Data.DataTable dtCost = MakeTableCombineEu2();
            System.Data.DataTable dt = Gettteu2();
            System.Data.DataTable dt1 = null;
            System.Data.DataTable dt2 = null;
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("?L????");
                return;
            }
            DataRow dr = null;
            for (int i = 0; i <= dt.Rows.Count - 1; i++)
            {
                DataRow dd = dt.Rows[i];
                dr = dtCost.NewRow();

                string gg = dd["code"].ToString();
                string SEQNO = dd["SEQNO"].ToString();
                string FDF = dd["???X"].ToString();
                dr["???X"] = "'" + FDF;
                dr["?????W??"] = dd["?????W??"].ToString();
                dr["?????W??2"] = dd["?????W??2"].ToString();
                dr["???J????"] = dd["???J????"].ToString();

                dr["?J?b???B"] = dd["?J?b???B"].ToString();
                dr["?????O"] = dd["?????O"].ToString();
                dr["???????B"] = dd["???????B"].ToString();
                dr["?q???O"] = dd["?q???O"].ToString();
                dr["??????"] = dd["??????"].ToString();
                dr["????????"] = dd["????????"].ToString();
                dr["?P??"] = dd["?P??"].ToString();
                dr["?~?U"] = dd["?~?U"].ToString();
                dr["?~???W??"] = dd["?~???W??"].ToString();
                string USD = dd["USD"].ToString();
                string RATE = dd["???v"].ToString();
                dr["???v"] = RATE;
                dr["NTD"] = dd["NTD"].ToString();

                dr["USD"] = USD;
                dt1 = Gettteu1(gg, SEQNO);
                StringBuilder sb = new StringBuilder();
                for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                {
                    DataRow dv = dt1.Rows[j];
                    string GH = dv["????"].ToString();
                    if (!String.IsNullOrEmpty(GH))
                    {
                        sb.Append(GH + "&");

                    }
                }
                if (!String.IsNullOrEmpty(sb.ToString()))
                {
                    sb.Remove(sb.Length - 1, 1);
                }
                string df = "";
                string docentry = sb.ToString();
                if (!String.IsNullOrEmpty(docentry))
                {
                    df = docentry + "/";
                }
                dr["?P??????"] = docentry;
                if (RATE == "1.000" || String.IsNullOrEmpty(RATE))
                {
                    dr["????"] = df + dr["?????W??"].ToString().Substring(0, 2);
                }
                else
                {
                    dr["????"] = df + dr["?????W??"].ToString().Substring(0, 2) + " US" + USD + "*" + RATE;

                }
                dt2 = Gettteu2(gg, SEQNO);

                StringBuilder sJ = new StringBuilder();
                for (int K = 0; K <= dt2.Rows.Count - 1; K++)
                {
                    DataRow dK = dt2.Rows[K];
                    string fg = dK["AR????"].ToString();
                    if (!String.IsNullOrEmpty(fg))
                    {
                        sJ.Append(fg + "&");
                    }

                }

                if (!String.IsNullOrEmpty(sJ.ToString()))
                {
                    sJ.Remove(sJ.Length - 1, 1);
                }

                dr["AR????"] = sJ.ToString();

                dtCost.Rows.Add(dr);

            }


            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACC\\TT2.xls";


            //Excel????????
            string ExcelTemplate = FileName;

            //???X??
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //???? Excel Report
            ExcelReport.ExcelReportOutput(dtCost, ExcelTemplate, OutPutFile, "N");
        }
        private void button22_Click(object sender, EventArgs e)
        {
            Execu();
        }





        private void button21_Click(object sender, EventArgs e)
        {

            DataRow drw = LCXX2.Rows[0];
            string aa = drw["filepath"].ToString();
            System.Diagnostics.Process.Start(aa);
        }

        private System.Data.DataTable GetOrderDataAP1(string TTCODE, string ID)
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT 'AR' ?`??,CAST(T0.DOCENTRY AS VARCHAR) ????,1+t1.vatprcnt/100 ?|?v,cast(t1.price as int) ?x??????,Convert(varchar(10),t0.docdate,112) ?L?b????,T0.U_IN_BSINV ?o?????X,");
            sb.Append("               T0.[Cardcode] ?????N?X,T0.[CardName] ?????W??,T1.ITEMCODE ???~?s??,Substring (T1.[ItemCode],2,8) ?~?W,CAST(T0.doctotal AS INT) ?x?????B ");
            sb.Append("              ,case T1.QUANTITY when 0 then 1 else CAST(T1.QUANTITY AS INT) end ???q,");
            sb.Append("              T10.ACCTCODE+' - '+T10.ACCTNAME ?o???`??,T0.COMMENTS ????,t9.u_acme_pay ????????,t1.u_acme_workday ?u?@????,t0.u_acme_paygui ?o?????B,CASE ISNULL(T8.PRICE,0) WHEN 0 THEN T1.u_acme_inv ELSE case t9.doccur when 'NTD' THEN T1.u_acme_inv ELSE CAST(T8.PRICE AS NVARCHAR) END END   ????????,T0.JRNLMEMO ?K?n,cast(T8.docentry as varchar) ?q?????X,t9.u_beneficiary ????????, ");
            sb.Append(" dbo.fun_CreditDate(T9.u_acme_pay,T0.CardCode,T0.DocDate) ?O??????,T0.u_in_bscls ?X?f???????O,T0.u_in_bsren ?X?f???????????X,T0.u_acme_shipto1 SHIPTO");
            sb.Append(" FROM OINV T0  ");
            sb.Append("              LEFT JOIN INV1 T1 ON T0.DocEntry = T1.DocEntry  ");
            sb.Append("              LEFT JOIN DLN1 T7 ON (T7.docentry=T1.baseentry AND T7.linenum=T1.baseline)");
            sb.Append("              LEFT JOIN RDR1 T8 ON (T8.docentry=T7.baseentry AND T8.linenum=T7.baseline)");
            sb.Append("              LEFT JOIN ORDR T9 ON (T8.docentry=T9.docentry )");
            sb.Append("              LEFT JOIN OACT T10 ON (T1.ACCTCODE=T10.ACCTCODE )");
            sb.Append("              where t0.docentry in (SELECT MEMO  FROM acmesqlsp.dbo.SATT2 WHERE  TTCODE=@TTCODE AND ID=@ID)  and t1.basetype='15' ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TTCODE", TTCODE));
            command.Parameters.Add(new SqlParameter("@ID", ID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "odln");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button11_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            textBox9.Text = "";
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void button11_Click_1(object sender, EventArgs e)
        {
            dataGridView7.DataSource = GETARPAY();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView7);
        }







    }
}


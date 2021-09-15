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
using Microsoft.VisualBasic.Devices;
namespace ACME
{

    public partial class LC : ACME.fmBase1
    {
        string strCn16 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn20 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP20;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn21 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP21;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string strCn22 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP22;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string NewFileName;
        System.Data.DataTable ddxx;
        System.Data.DataTable ddxx2;
        System.Data.DataTable dtCost=null;
        string PAY = "";
        public LC()
        {
            InitializeComponent();
        }

        private void account_LC2DataGridView_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            if (account_LC1DataGridView.SelectedRows.Count == 0 )
            {
                MessageBox.Show("�п�ܭn�R�P����");
                return;
            }
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            account_LCTableAdapter.Connection = MyConnection;
            account_LC1TableAdapter.Connection = MyConnection;
            account_LC2TableAdapter.Connection = MyConnection;
            account_LC3TableAdapter.Connection = MyConnection;
        }
        public override void AfterLoad()
        {
            button5.Enabled = true;
            button11.Enabled = true;


        }
        public override void AfterCancelEdit()
        {
            Control();
       
        }
        public override void AfterEdit()
        {
            oCCURTextBox.ReadOnly = true;
        }
        public override void query()
        {
            button5.Enabled = true;
            button11.Enabled = true;
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
        }
        private void Control()
        {

            lCTYPETextBox.ReadOnly = true;
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            button3.Enabled = true;
            button2.Enabled = true;
            button4.Enabled = true;
            button5.Enabled = true;
            button11.Enabled = true;
            button6.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
            button9.Enabled = true;
            button10.Enabled = true;
            comboBox1.Enabled = true;
            textBox5.ReadOnly = true;
            textBox4.ReadOnly = true;
            textBox3.ReadOnly = true;
            textBox6.ReadOnly = true;
            button12.Enabled = true;
            account_LC1DataGridView.Enabled = true;
            account_LC1DataGridView.ReadOnly = false;
            textBox5.ForeColor = Color.Red;
            button12.Enabled = true;
            button15.Enabled = true;
            oCCURTextBox.ReadOnly = true;
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
                mail.Account_LC.RejectChanges();
                mail.Account_LC1.RejectChanges();
                mail.Account_LC2.RejectChanges();
                mail.Account_LC3.RejectChanges();
            }
            catch { 
            }
            return true;
        }
        public override void SetDefaultValue()
        {
            oCCURTextBox.Text = "USD";
            string NumberName = "AL" + DateTime.Now.ToString("yyyyMMdd");
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);

            this.lCCODETextBox.Text = NumberName + AutoNum + "X";

            this.account_LCBindingSource.EndEdit();

            dONECheckBox.Checked = false;
           // lcAmtTextBox.Text = "0";
        }
        public override void SetInit()
        {

            MyBS = account_LCBindingSource;
            MyTableName = "Account_LC";
            MyIDFieldName = "LCCODE";
        }
        public override void FillData()
        {
            try
            {

                account_LCTableAdapter.Fill(mail.Account_LC, MyID);
                account_LC1TableAdapter.Fill(mail.Account_LC1, MyID);
                account_LC2TableAdapter.Fill(mail.Account_LC2, MyID);
                account_LC3TableAdapter.Fill(mail.Account_LC3, MyID);

                ddxx = GetMenu.Account_LCDownload(lCCODETextBox.Text,"1");
                if (ddxx.Rows.Count > 0)
                {
                    button5.Visible = true;
                }
                else
                {
                    button5.Visible = false;
                }

                ddxx2 = GetMenu.Account_LCDownload(lCCODETextBox.Text, "2");
                if (ddxx2.Rows.Count > 0)
                {
                    button11.Visible = true;
                }
                else
                {
                    button11.Visible = false;
                }


                decimal iTotal = 0;
                try
                {


                    int i = this.account_LC1DataGridView.Rows.Count - 1;
                    for (int iRecs = 0; iRecs <= i; iRecs++)
                    {

                        iTotal += Convert.ToDecimal(account_LC1DataGridView.Rows[iRecs].Cells["Amount"].Value);

                    }
                }
                catch (Exception ex)
                {
                }
                decimal iTotal2 = 0;
                try
                {
                    int i = this.account_LC2DataGridView.Rows.Count - 1;
                    for (int iRecs = 0; iRecs <= i; iRecs++)
                    {

                        iTotal2 += Convert.ToDecimal(account_LC2DataGridView.Rows[iRecs].Cells["Amountt"].Value);

                    }
                }
                catch (Exception ex)
                {
                }

                decimal F3 = Math.Round(iTotal, 2, MidpointRounding.AwayFromZero);
                decimal F4 = Math.Round(iTotal2, 2, MidpointRounding.AwayFromZero);
                decimal F5 = Math.Round(iTotal2 - iTotal, 2, MidpointRounding.AwayFromZero);
                decimal F6 = Math.Round(iTotal - iTotal2, 2, MidpointRounding.AwayFromZero);
                textBox3.Text = F3.ToString("#,##0.00");
                textBox4.Text = F4.ToString("#,##0.00");
                textBox5.Text = F5.ToString("#,##0.00");
                textBox6.Text = F6.ToString("#,##0.00");

        
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

                account_LCTableAdapter.Connection.Open();



                Validate();

                account_LCBindingSource.EndEdit();
                account_LC1BindingSource.EndEdit();
                account_LC2BindingSource.EndEdit();
                account_LC3BindingSource.EndEdit();


                tx = account_LCTableAdapter.Connection.BeginTransaction();



                SqlDataAdapter oWhsAdapter = GetAdapter(account_LCTableAdapter);
                oWhsAdapter.UpdateCommand.Transaction = tx;
                oWhsAdapter.InsertCommand.Transaction = tx;
                oWhsAdapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter oWhsAdapter1 = GetAdapter(account_LC1TableAdapter);
                oWhsAdapter1.UpdateCommand.Transaction = tx;
                oWhsAdapter1.InsertCommand.Transaction = tx;
                oWhsAdapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter oWhsAdapter2 = GetAdapter(account_LC2TableAdapter);
                oWhsAdapter2.UpdateCommand.Transaction = tx;
                oWhsAdapter2.InsertCommand.Transaction = tx;
                oWhsAdapter2.DeleteCommand.Transaction = tx;

                SqlDataAdapter oWhsAdapter3 = GetAdapter(account_LC3TableAdapter);
                oWhsAdapter3.UpdateCommand.Transaction = tx;
                oWhsAdapter3.InsertCommand.Transaction = tx;
                oWhsAdapter3.DeleteCommand.Transaction = tx;

                account_LCTableAdapter.Update(mail.Account_LC);
                mail.Account_LC.AcceptChanges();
                account_LC1TableAdapter.Update(mail.Account_LC1);
                mail.Account_LC1.AcceptChanges();
                account_LC2TableAdapter.Update(mail.Account_LC2);
                mail.Account_LC2.AcceptChanges();
                account_LC3TableAdapter.Update(mail.Account_LC3);
                mail.Account_LC3.AcceptChanges();

                this.MyID = this.lCCODETextBox.Text;
                tx.Commit();


                UpdateData = true;
            }
            catch (Exception ex)
            {
                if (tx != null)
                {

                    tx.Rollback();

                }

                MessageBox.Show(ex.Message, "��s���~", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.account_LCTableAdapter.Connection.Close();



            }
            return UpdateData;



        }

        private void account_LC1DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["Seqno"].Value = util.GetSeqNo(2, account_LC1DataGridView);
            e.Row.Cells["Quantity"].Value = 0;
            e.Row.Cells["Price"].Value = 0;
            
        }
        public void OINV()
        {
            

                    System.Data.DataTable G2 = GETOINV2();
                    if (G2.Rows.Count > 0)
                    {
                        for (int i = 0; i <= G2.Rows.Count - 1; i++)
                        {
                            string ID = G2.Rows[i]["ID"].ToString();
                            string docentry = G2.Rows[i]["docentry"].ToString();
                            string linenum = G2.Rows[i]["linenum"].ToString();
                            string QTY = G2.Rows[i]["QTY"].ToString();


                            System.Data.DataTable G1 = GETOINV(docentry, linenum);
                            if (G1.Rows.Count > 0)
                            {
                                UpdateOINV(ID, G1.Rows[0][0].ToString());
                            }
                            else
                            {
                                System.Data.DataTable G11 = GETOINV11(docentry, linenum,QTY);
                                if (G11.Rows.Count > 0)
                                {
                                    UpdateOINV(ID, G11.Rows[0][0].ToString());
                                }
                            
                            }
                        }

                    }

        }
      

        private void account_LC1DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (account_LC1DataGridView.Columns[e.ColumnIndex].Name == "Quantity" ||
                          account_LC1DataGridView.Columns[e.ColumnIndex].Name == "Price")
                {

                    decimal iQuantity = 0;
                    decimal iUnitPrice = 0;
              
                    iQuantity = Convert.ToInt32(this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Quantity"].Value);
                    iUnitPrice = Convert.ToDecimal(this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Price"].Value);
                    this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Amount"].Value = (iQuantity * iUnitPrice).ToString();

                }

                if (account_LC1DataGridView.Columns[e.ColumnIndex].Name == "docentry" ||
           account_LC1DataGridView.Columns[e.ColumnIndex].Name == "linenum")
                {
                    string docentry = this.account_LC1DataGridView.Rows[e.RowIndex].Cells["docentry"].Value.ToString();
                    string linenum = this.account_LC1DataGridView.Rows[e.RowIndex].Cells["linenum"].Value.ToString();

                    if (globals.DBNAME == "�i����")
                    {
                        if (cOMPANYTextBox.Text == "CHOICE" || cOMPANYTextBox.Text == "TOP" || cOMPANYTextBox.Text == "Infinite")
                        {
                            System.Data.DataTable d2 = GetQTYINF(docentry, linenum, cOMPANYTextBox.Text);

                            if (d2.Rows.Count > 0)
                            {

                                this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Quantity"].Value = Convert.ToInt32(d2.Rows[0][0].ToString());
                                this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Price"].Value = Convert.ToDecimal(d2.Rows[0][1].ToString());
                                this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Model"].Value = d2.Rows[0][2].ToString();

                            }
                        }
                        else
                        {
                            System.Data.DataTable d1 = GetQTY(docentry, linenum);

                            if (d1.Rows.Count > 0)
                            {

                                this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Quantity"].Value = Convert.ToInt32(d1.Rows[0][0].ToString());
                                this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Price"].Value = Convert.ToDecimal(d1.Rows[0][1].ToString());
                                this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Model"].Value = d1.Rows[0][2].ToString();

                            }
                        }
                    }
                    else
                    {

                        System.Data.DataTable d2 = GetQTYINF(docentry, linenum, globals.DBNAME);



                        if (d2.Rows.Count > 0)
                        {

                            this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Quantity"].Value = Convert.ToInt32(d2.Rows[0][0].ToString());
                            this.account_LC1DataGridView.Rows[e.RowIndex].Cells["Price"].Value = Convert.ToDecimal(d2.Rows[0][1].ToString());
 

                        }
                    }
                }

            }
            catch (Exception ex)
            {
             
            }
        }

        private void account_LC2DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
          
            try
            {
                if (account_LC2DataGridView.Columns[e.ColumnIndex].Name == "Quantityy" ||
                          account_LC2DataGridView.Columns[e.ColumnIndex].Name == "Price2")
                {

                    decimal iQuantity = 0;
                    decimal iUnitPrice = 0;

                    iQuantity = Convert.ToInt32(this.account_LC2DataGridView.Rows[e.RowIndex].Cells["Quantityy"].Value);
                    iUnitPrice = Convert.ToDecimal(this.account_LC2DataGridView.Rows[e.RowIndex].Cells["Price2"].Value);
                    this.account_LC2DataGridView.Rows[e.RowIndex].Cells["Amountt"].Value = (iQuantity * iUnitPrice).ToString();

                }
             
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void account_LC2DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            try
            {
                e.Row.Cells["dataGridViewTextBoxColumn1"].Value = account_LC1DataGridView.SelectedRows[0].Cells["Seqno"].Value.ToString();
                e.Row.Cells["dataGridViewTextBoxColumn2"].Value = account_LC1DataGridView.SelectedRows[0].Cells["Model"].Value.ToString();
                e.Row.Cells["Quantityy"].Value = account_LC1DataGridView.SelectedRows[0].Cells["Quantity"].Value.ToString();
                e.Row.Cells["Price2"].Value = account_LC1DataGridView.SelectedRows[0].Cells["Price"].Value.ToString();
                e.Row.Cells["Amountt"].Value = account_LC1DataGridView.SelectedRows[0].Cells["Amount"].Value.ToString();

            }
            catch (Exception ex)
            {
            }
        }
        public override void AfterEndEdit()
        {

            try
            {

                UpdateLCTYPE();

                System.Data.DataTable dt1 = GetTT(lCCODETextBox.Text);
         
                UpdateLC11(lCCODETextBox.Text);
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow row = dt1.Rows[i];
                    string seqno = row["seqno"].ToString();
                    decimal ���B = Convert.ToDecimal(row["���B"]);
                    int �ƶq = Convert.ToInt32(row["�ƶq"]);

                    UpdateLC1(���B, seqno,�ƶq, lCCODETextBox.Text);
                        
      
                    }
              
                    account_LCTableAdapter.Fill(mail.Account_LC, MyID);
                }
         
            
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void UpdateLCTYPE()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE Account_LC SET   lCTYPE='AT SIGHT' WHERE DRAFT ='�Y��' AND ISNULL(lCTYPE,'')='' ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


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
        private void UpdateOINV(string ID, string OINV)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE Account_LC1 SET OINV=@OINV WHERE ID=@ID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@OINV", OINV));
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
        private void UpdateLC1(decimal amount, string seqno, int quantity, string lccode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update Account_LC1 set quantity1=quantity-@quantity,amount1=amount-@amount where seqno=@seqno and lccode=@lccode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@quantity", quantity));
            command.Parameters.Add(new SqlParameter("@amount", amount));
            command.Parameters.Add(new SqlParameter("@seqno", seqno));
            command.Parameters.Add(new SqlParameter("@lccode", lccode));
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
        private void UpdateLC11(string lccode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update Account_LC1 set quantity1=quantity,amount1=amount where  lccode=@lccode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@lccode", lccode));
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


        public static System.Data.DataTable GetTT(string LCCODE)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "select seqno,sum(quantity) �ƶq,sum(amount) ���B from Account_LC2  where LCCODE=@LCCODE GROUP BY seqno HAVING sum(amount)  <> 0 ";

            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LCCODE", LCCODE));
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
        public static System.Data.DataTable GETOINV(string DOCENTRY,string LINENUM)
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.shipConnection;
            sb.Append(" SELECT T1.DOCENTRY FROM INV1 T1 ");
            sb.Append("  left join dln1 t4 on (t1.baseentry=T4.docentry and  t1.baseline=t4.linenum  and t1.basetype='15')");
            sb.Append(" left join rdr1 t5 on (t4.baseentry=T5.docentry and  t4.baseline=t5.linenum  and t5.targettype='15')");
            sb.Append(" WHERE T5.DOCENTRY=@DOCENTRY AND T5.LINENUM =@LINENUM");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LineNum", LINENUM));
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
        public static System.Data.DataTable GETOINV11(string DOCENTRY, string LINENUM,string QTY)
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.shipConnection;
            sb.Append(" SELECT T1.DOCENTRY FROM INV1 T1  ");
            sb.Append(" left join rdr1 t5 on (t1.baseentry=T5.docentry and  t1.baseline=t5.linenum  and t5.targettype='13') ");
            sb.Append(" WHERE T5.DOCENTRY=@DOCENTRY AND T5.LINENUM =@LINENUM and t1.quantity=@QTY ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LineNum", LINENUM));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
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
        public static System.Data.DataTable GETOINV2()
        {
            StringBuilder sb = new StringBuilder();
            SqlConnection MyConnection = globals.Connection;
            sb.Append("SELECT ID,B.docentry,B.linenum,CAST(B.QUANTITY AS INT) QTY  FROM Account_LC A  left join Account_LC1 b on (a.LCCODE=b.LCCODE) WHERE ISNULL(OINV,'') ='' AND A.COMPANY ='ACME' AND B.docentry<>'' AND CAST(B.linenum AS VARCHAR)<>''   ");
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
        private void account_LC2DataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex >= account_LC2DataGridView.Rows.Count)
                return;
            try
            {

                if (account_LC1DataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow dgr = account_LC2DataGridView.Rows[e.RowIndex];
                    string da = account_LC1DataGridView.SelectedRows[0].Cells["Seqno"].Value.ToString();
                    string dd = dgr.Cells["dataGridViewTextBoxColumn1"].Value.ToString();
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

 

        private void button1_Click(object sender, EventArgs e)
        {
                 object[] LookupValues = null;

                 if (globals.DBNAME == "�t��")
                 {
                     LookupValues = GetMenu.GetCHICUST14();
                 }
                 else
                 {
                     LookupValues = GetMenu.GetMenuList();
                 }


            if (LookupValues != null)
            {
                cARDCODETextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAMETextBox.Text = Convert.ToString(LookupValues[1]);

            }
        }


        private void GetLC()
        {
            try
            {

                System.Data.DataTable dt = null;

                if (PAY == "A1")
                {

                    dt = GetLC1();
                }
                else if (PAY == "A2")
                {
                    dt = GetLCARDPAY();
                }
           
                System.Data.DataTable dt1 = null;

                dtCost = MakeTableCombine();
                DataRow dr = null;
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                   string  seqno = dt.Rows[i]["seqno"].ToString();
                   string  CSSCODE = dt.Rows[i]["LCCODE"].ToString();

                    dt1 = GetLC2(CSSCODE, seqno);
                    dr = dtCost.NewRow();
    

                    dr["�}����"] = dt.Rows[i]["�}����"].ToString();
                    dr["�}����"] = dt.Rows[i]["�}����"].ToString();
                    dr["�Y��"] = dt.Rows[i]["�Y��"].ToString();
                    dr["�}���H"] = dt.Rows[i]["�}���H"].ToString();
                    dr["LCNO"] = dt.Rows[i]["LC"].ToString();
                    dr["�~�W"] = dt.Rows[i]["�~�W"].ToString();
                    dr["�Ƶ�"] = dt.Rows[i]["�Ƶ�"].ToString();
                    dr["�}���H"] = dt.Rows[i]["�}���H"].ToString();
                    dr["LCNO"] = dt.Rows[i]["LC"].ToString();
                    dr["�~�W"] = dt.Rows[i]["�~�W"].ToString();
                    dr["�ƶq"] = Convert.ToInt32(dt.Rows[i]["�ƶq"].ToString());
                    dr["���B"] = Convert.ToDecimal(dt.Rows[i]["���B"].ToString());
                    dr["�l�B�ƶq"] = Convert.ToInt32(dt.Rows[i]["�l�B�ƶq"].ToString());
                    dr["�l�B"] = Convert.ToDecimal(dt.Rows[i]["�l�B"].ToString());
                    dr["�w�R�ƶq"] = Convert.ToInt32(dt.Rows[i]["�w�R�ƶq"].ToString());
                    dr["�w�R���B"] = Convert.ToDecimal(dt.Rows[i]["�w�R���B"].ToString());
                    dr["EXPIRY"] = dt.Rows[i]["EXPIRY"].ToString();
                    dr["�̫�˲��"] = dt.Rows[i]["�̫�˲��"].ToString();
                    dr["�~�U"] = dt.Rows[i]["�~�U"].ToString();
                    StringBuilder sb = new StringBuilder();
                    StringBuilder sd = new StringBuilder();
                    StringBuilder sh = new StringBuilder();

                    StringBuilder G1 = new StringBuilder();
                    StringBuilder G2 = new StringBuilder();
                    StringBuilder G3 = new StringBuilder();
                    StringBuilder G4 = new StringBuilder();
                    StringBuilder G5 = new StringBuilder();
                    
                    for (int j = 0; j <= dt1.Rows.Count - 1; j++)
                    {
                        DataRow dv = dt1.Rows[j];

                        string se = (Convert.ToInt32(dv["�R�P���B"])).ToString();

                        sb.Append(dv["�R�P�ƶq"].ToString() + "/");
                        sd.Append(se + "/");

                        sh.Append(dv["��פ��"].ToString() + "/");
                        string F1 = dv["��צ�"].ToString();
                        string F2 = dv["���O"].ToString();
                        string F3 = dv["��פ�"].ToString();
                        string F4 = dv["�w�p�ӧI��"].ToString();
                        string F5 = dv["��ڤJ�b��"].ToString();

                        if (!String.IsNullOrEmpty(F1))
                        {
                            G1.Append(F1 + "/");
                        }
                        if (!String.IsNullOrEmpty(F2))
                        {
                            G2.Append(F2 + "/");
                        }
                        if (!String.IsNullOrEmpty(F3))
                        {
                            G3.Append(F3 + "/");
                        }
                        if (!String.IsNullOrEmpty(F4))
                        {
                            G4.Append(F4 + "/");
                        }
                        if (!String.IsNullOrEmpty(F5))
                        {
                            G5.Append(F5 + "/");
                        }

                    }
                    if (!String.IsNullOrEmpty(sb.ToString()))
                    {
                        sb.Remove(sb.Length - 1, 1);
                    }
                    if (!String.IsNullOrEmpty(sd.ToString()))
                    {
                        sd.Remove(sd.Length - 1, 1);
                    }
                    if (!String.IsNullOrEmpty(sh.ToString()))
                    {
                        sh.Remove(sh.Length - 1, 1);
                    }

                    if (!String.IsNullOrEmpty(G1.ToString()))
                    {
                        G1.Remove(G1.Length - 1, 1);
                    }
                    if (!String.IsNullOrEmpty(G2.ToString()))
                    {
                        G2.Remove(G2.Length - 1, 1);
                    }
                    if (!String.IsNullOrEmpty(G3.ToString()))
                    {
                        G3.Remove(G3.Length - 1, 1);
                    }
                    if (!String.IsNullOrEmpty(G4.ToString()))
                    {
                        G4.Remove(G4.Length - 1, 1);
                    }
                    if (!String.IsNullOrEmpty(G5.ToString()))
                    {
                        G5.Remove(G5.Length - 1, 1);
                    }
                 
                    dr["�R�P�ƶq"] = sb.ToString();
                    dr["�R�P���B"] = sd.ToString();
                    dr["��פ��"] = sh.ToString();

                    dr["��צ�"] = G1.ToString();
                    dr["���O"] = G2.ToString();
                    dr["��פ�"] = G3.ToString();
                    dr["�w�p�ӧI��"] = G4.ToString();
                    dr["��ڤJ�b��"] = G5.ToString();
                    

                    dtCost.Rows.Add(dr);
                }
        
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        private System.Data.DataTable GetLCARDPAY()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("            select  APPLICANT,SUBSTRING(A.iSSUEDATE,1,4)+'/'+SUBSTRING(A.iSSUEDATE,5,2)+'/'+SUBSTRING(A.iSSUEDATE,7,2) �}����,A.dRAWEE �}����,A.dRAFTDay �Y��,A.LCCODE,b.seqno,substring(a.cardname,0,5) �}���H,lcno LC,b.MODEL �~�W,ISNULL(b.QUANTITY,0) �ƶq,ISNULL(b.AMOUNT,0) ���B,ISNULL(b.QUANTITY1,0) �l�B�ƶq,ISNULL(b.AMOUNT1,0) �l�B,ISNULL(b.QUANTITY,0)-ISNULL(b.QUANTITY1,0) �w�R�ƶq,ISNULL(b.AMOUNT,0)-ISNULL(b.AMOUNT1,0) �w�R���B,EXPIRY EXPIRY,A.ShipDate �̫�˲��,a.mEMO �Ƶ�,SA �~�U from Account_LC a    ");
            sb.Append("            left join Account_LC1 b on (a.LCCODE=b.LCCODE)");
            sb.Append(" LEFT JOIN ACMESQL02.DBO.RDR1 T2 ON (B.DOCENTRY=T2.DOCENTRY AND B.LINENUM=T2.LINENUM)");
            sb.Append("            left join ACMESQL02.DBO.dln1 T3 on (T3.baseentry=T2.docentry and  T3.baseline=t2.linenum  and T3.basetype='17')");
            sb.Append("            left join ACMESQL02.DBO.INV1 T4 on (T4.baseentry=T3.docentry and  T4.baseline=T3.linenum  and T4.basetype='15')");
            sb.Append("             WHERE t4.docentry=@aa order by APPLICANT ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox7.Text));
    

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

        private System.Data.DataTable GetLC1()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" select  APPLICANT,SUBSTRING(A.iSSUEDATE,1,4)+'/'+SUBSTRING(A.iSSUEDATE,5,2)+'/'+SUBSTRING(A.iSSUEDATE,7,2) �}����,A.dRAWEE �}����,A.dRAFTDay �Y��,A.LCCODE,b.seqno,substring(a.cardname,0,5) �}���H,lcno LC,b.MODEL �~�W,ISNULL(b.QUANTITY,0) �ƶq,ISNULL(b.AMOUNT,0) ���B,ISNULL(b.QUANTITY1,0) �l�B�ƶq,ISNULL(b.AMOUNT1,0) �l�B,ISNULL(b.QUANTITY,0)-ISNULL(b.QUANTITY1,0) �w�R�ƶq,ISNULL(b.AMOUNT,0)-ISNULL(b.AMOUNT1,0) �w�R���B,EXPIRY EXPIRY,ShipDate �̫�˲��,a.mEMO �Ƶ�,SA �~�U from Account_LC a    ");
            sb.Append(" left join Account_LC1 b on (a.LCCODE=b.LCCODE)");
            sb.Append("  WHERE A.iSSUEDATE between @aa and @bb order by APPLICANT ");
            
         
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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
        private System.Data.DataTable GetLC2(string CSSCODE, string seqno)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  quantity �R�P�ƶq,amount �R�P���B,LCTTPE2 ��פ��,Bank ��צ�,LCTTPE	���O,LCDATE	��פ�,INDATE �w�p�ӧI��,REDATE ��ڤJ�b�� FROM Account_LC2 ");
            sb.Append("  WHERE LCCODE=@LCCODE AND seqno=@seqno ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@LCCODE", CSSCODE));
            command.Parameters.Add(new SqlParameter("@seqno", seqno));

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

        private System.Data.DataTable GetLC3()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("        select applicant �}���H,lcno LC,AMOUNT ���B,LCDATE ��פ�,INDATE �w�p�ӧI��,REDATE ��ڤJ�b��,b.MEMO LC�����O�� from Account_LC a  ");
            sb.Append("              left join Account_LC2 b on (a.LCCODE=b.LCCODE)  ");
            sb.Append("  WHERE A.iSSUEDATE between @aa and @bb order by APPLICANT ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@bb", textBox2.Text));

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

        private System.Data.DataTable GetOINVDATE(string DOCENTRY)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("   SELECT Convert(varchar(10),DOCDATE,112) DOCDATE FROM OINV WHERE DOCENTRY=@DOCENTRY ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));


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
        private void button1_Click_1(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetMenuList();

            if (LookupValues != null)
            {
                cARDCODETextBox.Text = Convert.ToString(LookupValues[0]);
                cARDNAMETextBox.Text = Convert.ToString(LookupValues[1]);

            }
        }

        

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            dt.Columns.Add("�}����", typeof(string));
            dt.Columns.Add("�}����", typeof(string));
            dt.Columns.Add("�Y��", typeof(string));
            dt.Columns.Add("��פ��", typeof(string));
            dt.Columns.Add("�}���H", typeof(string));
            dt.Columns.Add("LCNO", typeof(string));
            dt.Columns.Add("�~�W", typeof(string));
            dt.Columns.Add("�ƶq", typeof(Int32));
            dt.Columns.Add("���B", typeof(decimal));
            dt.Columns.Add("�R�P�ƶq", typeof(string));
            dt.Columns.Add("�R�P���B", typeof(string));
            dt.Columns.Add("�l�B�ƶq", typeof(Int32));
            dt.Columns.Add("�l�B", typeof(decimal));
            dt.Columns.Add("�Ƶ�", typeof(string));
            dt.Columns.Add("�w�R�ƶq", typeof(Int32));
            dt.Columns.Add("�w�R���B", typeof(decimal));
            dt.Columns.Add("EXPIRY", typeof(string));
            dt.Columns.Add("�̫�˲��", typeof(string));
            dt.Columns.Add("�~�U", typeof(string));
            dt.Columns.Add("��צ�", typeof(string));
            dt.Columns.Add("���O", typeof(string));
            dt.Columns.Add("��פ�", typeof(string));
            dt.Columns.Add("�w�p�ӧI��", typeof(string));
            dt.Columns.Add("��ڤJ�b��", typeof(string));
            return dt;
        }

        private void button3_Click(object sender, EventArgs e)
        {

            System.Data.DataTable FF = GetLC3();
            ACME.FormTT2 frm = new ACME.FormTT2();
            frm.dt = FF;
            frm.ShowDialog();
        }

        private void LC_Load_1(object sender, EventArgs e)
        {
            Control();
            OINV();
            textBox1.Text = GetMenu.DFirst();
            textBox2.Text = GetMenu.DLast();

            ddxx = GetMenu.Account_LCDownload(lCCODETextBox.Text,"1");
            if (ddxx.Rows.Count > 0)
            {
                button5.Visible = true;
            }
            else
            {
                button5.Visible = false;
            }


            ddxx2 = GetMenu.Account_LCDownload(lCCODETextBox.Text, "2");
            if (ddxx2.Rows.Count > 0)
            {
                button11.Visible = true;
            }
            else
            {
                button11.Visible = false;
            }
            comboBox1.Text = "ACME";

            account_LC2DataGridView.ReadOnly = true;



        }

        private void button4_Click(object sender, EventArgs e)
        {


            string server = "//acmesrv01//SAP_Share//Rma//";
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            string filename = Path.GetFileName(opdf.FileName);
            if (result == DialogResult.OK)
            {
                DelDownload(lCCODETextBox.Text,"1");
                string file = opdf.FileName;
                bool F1 = getrma.UploadFile(file, server, false);
                if (F1 == false)
                {
                    return;
                }
                AddDownload(lCCODETextBox.Text, filename, @"\\acmesrv01\SAP_Share\Rma\" + filename, "1");
                MessageBox.Show("�W�Ǧ��\");
                button5.Visible = true;

            }
            
        }
        public void AddDownload(string LCID,string Download, string FilePath,string ID2)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into Account_LCDownload(LCID,Download,FilePath,ID2) values(@LCID,@Download,@FilePath,@ID2)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LCID", LCID));
            command.Parameters.Add(new SqlParameter("@Download", Download));
            command.Parameters.Add(new SqlParameter("@FilePath", FilePath));
            command.Parameters.Add(new SqlParameter("@ID2", ID2));
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
        public void DelDownload(string LCID, string ID2)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("delete Account_LCDownload where LCID=@LCID AND ID2=@ID2 ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@LCID", LCID));
            command.Parameters.Add(new SqlParameter("@ID2", ID2));
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

            try
            {
                DataRow drw = ddxx.Rows[0];
                string aa = drw["filepath"].ToString();
                System.Diagnostics.Process.Start(aa);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                Cursor = Cursors.WaitCursor;

                try
                {
                    string FileName = openFileDialog1.FileName;

                   
                        GetExcelProduct(FileName, 1, 1, 11, 3);
                                        
                    MessageBox.Show("�����ɮ�->" + NewFileName);
                }
                finally
                {
                    Cursor = Cursors.Default;
                }

            }
        }

        private void GetExcelProduct(string ExcelFile, int a, int b, int c, int d)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = true;

            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            //���o  Worksheet
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);


            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;


            Microsoft.Office.Interop.Excel.Range range = null;


            try
            {

                string sTemp = string.Empty;
                string FieldValue = string.Empty;


                for (int iRecord = iRowCnt; iRecord >= d; iRecord--)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, a]);
                    range.Select();
                    sTemp = (string)range.Text;

                    if (sTemp == "" || sTemp == "�}���H")
                    {
                        for (int i = b; i <= c; i++)
                        {
                            range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, i]);
                            range.Select();
                            range.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                        }
                    }
               


                }

            }
            finally
            {

                NewFileName = Path.GetDirectoryName(ExcelFile) + "\\" +
               "Acme_" + Path.GetFileNameWithoutExtension(ExcelFile) + ".xls";
                //GetFileName(ExcelFile);
                //   MessageBox.Show(NewFileName);

                try
                {
                    excelSheet.SaveAs(NewFileName, XlFileFormat.xlExcel9795, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                    //  excelSheet.SaveAs(NewFileName, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
                }
                catch
                {
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
                //�i�H�N Excel.exe �M��
                System.GC.WaitForPendingFinalizers();
                // MessageBox.Show("���ͤ@���ɮ�->" + NewFileName);



            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            string server = "//acmesrv01//SAP_Share//Rma//";
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            string filename = Path.GetFileName(opdf.FileName);
            if (result == DialogResult.OK)
            {
                DelDownload(lCCODETextBox.Text, "2");
                string file = opdf.FileName;

                bool F1 = getrma.UploadFile(file, server, false);
                if (F1 == false)
                {
                    return;
                }
                AddDownload(lCCODETextBox.Text, filename, @"\\acmesrv01\SAP_Share\Rma\" + filename, "2");
                MessageBox.Show("�W�Ǧ��\");
                button11.Visible = true;

            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetTemp5();
            ExcelReport.GridViewToExcel(dataGridView1);

        }
        System.Data.DataTable GetTemp5()
        {
            DateTime before1month = DateTime.Now.AddMonths(1);
            string ee = before1month.ToString("yyyyMMdd");
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" select applicant �}���H,lcno LC,shipDate �˲��,eXPIRY �����,OrderDate �q��̫��f��,SUM(T1.QUANTITY) �ƶq");
            sb.Append(" ,SUM(T1.AMOUNT) ���B,SUM(QUANTITY1) �l�B�ƶq,SUM(AMOUNT1) �l�B,MEMO �Ƶ� from dbo.Account_LC t0 ");
            sb.Append(" left join dbo.Account_LC1 t1 on (t0.lccode=t1.lccode)");
            sb.Append(" where amount1 <> 0  and eXPIRY <=  '" + ee + "' GROUP BY applicant ,lcno ,shipDate ,eXPIRY,OrderDate,MEMO order by eXPIRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

 

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetTemp6();
            ExcelReport.GridViewToExcel(dataGridView1);
        }
        System.Data.DataTable GetTemp6()
        {

            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT APPLICANT �Ȥ�W��,LCNO �H�Ϊ����X,LCDATE ��פ�,INDATE �w�p�ӧI��,REDATE ��ڤJ�b�� FROM ACCOUNT_LC2 T0");
            sb.Append(" LEFT JOIN ACCOUNT_LC T1 ON(T0.LCCODE=T1.LCCODE)");
            sb.Append(" WHERE ISNULL(LCDATE,'') ='' OR ISNULL(INDATE,'') ='' OR ISNULL(REDATE,'')='' ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;



            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        System.Data.DataTable GetTemp66()
        {
            
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("         select   COMPANY,APPLICANT,SUBSTRING(A.iSSUEDATE,1,4)+'/'+SUBSTRING(A.iSSUEDATE,5,2)+'/'+SUBSTRING(A.iSSUEDATE,7,2) �}����,A.dRAWEE �}����,A.dRAFTDay �Y��,substring(a.cardname,0,5) �}���H,lcno LC,b.MODEL �~�W,b.QUANTITY ��׼ƶq,b.AMOUNT ��ת��B,Bank ��צ�,LCTTPE ���O,LCTTPE2 ��פ��,LCDATE ��פ� ");
            sb.Append("               ,INDATE �w�p�ӧI��,REDATE ��ڤJ�b��,EXPIRY EXPIRY,ShipDate �̫�˲��,b.meMo LC�����O��,a.mEMO �Ƶ�,C.DOCENTRY �q�渹�X from Account_LC a    ");
            sb.Append("             left join Account_LC2 b on (a.LCCODE=b.LCCODE) ");
            sb.Append("     left join Account_LC1 C on (B.LCCODE=C.LCCODE AND B.SEQNO=C.SEQNO) ");
            sb.Append("             WHERE  1=1  ");
         if (comboBox1.Text != "ALL")
         {
             sb.Append("  AND  cOMPANY=@CC ");
         }
         sb.Append("  AND b.lcDATE between @AA and @BB order by COMPANY,APPLICANT");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@AA", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@BB", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@CC", comboBox1.Text));


            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        System.Data.DataTable GetQTY(string DOCENTRY, string LINENUM)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CAST(QUANTITY AS INT) �ƶq,PRICE ���,Substring ([ItemCode],2,8)+' V.'+Substring([ItemCode],12,1) MODEL FROM RDR1  WHERE DOCENTRY=@DOCENTRY AND CAST(LINENUM AS VARCHAR)=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));



            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }

        System.Data.DataTable GetQTYINF(string DOCENTRY, string LINENUM,string COMPANNY)
        {

            SqlConnection connection = null;

            if (COMPANNY == "�t��")
            {
                connection = new SqlConnection(strCn16);
            }
            else if (COMPANNY == "CHOICE")
            {
                connection = new SqlConnection(strCn21);
            }
            else if (COMPANNY == "TOP")
            {
                connection = new SqlConnection(strCn20);
            }
            else if (COMPANNY == "Infinite")
            {
                connection = new SqlConnection(strCn22);
            }
            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT CAST(QUANTITY AS INT) �ƶq,PRICE ���,Substring (ProdID,2,8)+' V.'+Substring(ProdID,12,1) FROM OrdBillSub  WHERE BillNO=@DOCENTRY AND CAST(SerNO AS VARCHAR)=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));



            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "wh_main");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["wh_main"];
        }
        private void dRAFTComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (dRAFTTextBox.Text == "����")
            {
                dRAFTDayTextBox.Text = "30 days";
            }
            if (dRAFTTextBox.Text == "�Y��")
            {
                lCTYPETextBox.Text = "AT SIGHT";
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            System.Data.DataTable dt = GetLC1();
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("�S�����");
                return;
            }
            PAY = "A1";
            GetLC();

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACLC.xls";

        
                //���o Excel ���
            System.Data.DataTable   OrderData = dtCost;
            
          
            //Excel���˪���
            string ExcelTemplate = FileName;

            //��X��
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //���� Excel Report
            ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetTemp66();
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                DataRow drw = ddxx2.Rows[0];
                string aa = drw["filepath"].ToString();
                System.Diagnostics.Process.Start(aa);
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

                this.account_LC1BindingSource.EndEdit();
                this.account_LC1TableAdapter.Update(mail.Account_LC1);
                mail.Account_LC1.AcceptChanges();
                MessageBox.Show("�x�s���\");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void account_LC1DataGridView_MouseEnter(object sender, EventArgs e)
        {

            if (account_LC1DataGridView.Rows.Count > 1)
            {
                if (globals.GroupID.ToString().Trim() == "SA")
                {
                    account_LC1DataGridView.Columns["Model"].ReadOnly = true;
                    account_LC1DataGridView.Columns["Quantity"].ReadOnly = true;
                    account_LC1DataGridView.Columns["Price"].ReadOnly = true;
                    account_LC1DataGridView.Columns["Amount"].ReadOnly = true;
                    account_LC1DataGridView.Columns["Quantity1"].ReadOnly = true;
                    account_LC1DataGridView.Columns["Amount1"].ReadOnly = true;
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {

                if (account_LC1DataGridView.SelectedRows.Count == 0)
                {
                    MessageBox.Show("�п�ܦC���");

                    return;
                }

                System.Data.DataTable dt2 = mail.Account_LC2;
    
         

                int i = account_LC1DataGridView.SelectedRows.Count - 1;
        
                for (int iRecs = i; iRecs >= 0; iRecs--)
                {
                    DataRow drw2 = dt2.NewRow();
                    //OINV1
                    string SEQNO = account_LC1DataGridView.SelectedRows[iRecs].Cells["Seqno"].Value.ToString();
                    string Model = account_LC1DataGridView.SelectedRows[iRecs].Cells["Model"].Value.ToString();
                    string Quantity = account_LC1DataGridView.SelectedRows[iRecs].Cells["Quantity"].Value.ToString();
                    string Price = account_LC1DataGridView.SelectedRows[iRecs].Cells["Price"].Value.ToString();
                    string Amount = account_LC1DataGridView.SelectedRows[iRecs].Cells["Amount"].Value.ToString();
                    string OINV1 = account_LC1DataGridView.SelectedRows[iRecs].Cells["OINV1"].Value.ToString();


                drw2["LCCODE"] = lCCODETextBox.Text;
                drw2["Seqno"] = SEQNO;
                drw2["Model"] = Model;
                drw2["Quantity"] = Quantity;
                drw2["Price"] = Price;
                drw2["Amount"] = Amount;
                if (globals.DBNAME != "�t��")
                {
                    System.Data.DataTable GH1 = GetOINVDATE(OINV1);
                    if (GH1.Rows.Count > 0)
                    {
                        drw2["DDATE"] = GH1.Rows[0][0].ToString();

                    }
                }
                dt2.Rows.Add(drw2);
              

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            this.account_LC1BindingSource.EndEdit();
        }
        public override void STOP()
        {
            if (lCNOTextBox.Text == "")
            {
                MessageBox.Show("�п�JLC NO");
                this.SSTOPID = "1";
                lCNOTextBox.Focus();
                return;
            }
   
        }

        private void account_LC2DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            string d1 = account_LC2DataGridView.Rows[0].Cells["bank"].Value.ToString();
            string d2 = account_LC2DataGridView.Rows[0].Cells["type"].Value.ToString();
            string d3 = account_LC2DataGridView.Rows[0].Cells["LCTTPE2"].Value.ToString();
            string d4 = account_LC2DataGridView.Rows[0].Cells["LDATE"].Value.ToString();
            DataGridViewRow row;
            for (int i = account_LC2DataGridView.Rows.Count - 2; i >= 0; i--)
            {
                row = account_LC2DataGridView.Rows[i];

                row.Cells["bank"].Value = d1;
                row.Cells["type"].Value = d2;
                row.Cells["LCTTPE2"].Value = d3;
                row.Cells["LDATE"].Value = d4;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            System.Data.DataTable  dt = GetLCARDPAY();
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("�S�����");
                return;
            }
            PAY = "A2";
            GetLC();

          
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\ACLC.xls";


            //���o Excel ���
            System.Data.DataTable OrderData = dtCost;


            //Excel���˪���
            string ExcelTemplate = FileName;

            //��X��
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //���� Excel Report
            ExcelReport.ExcelReportOutput(OrderData, ExcelTemplate, OutPutFile, "N");
        }

      

        private void button16_Click(object sender, EventArgs e)
        {
            string NumberName = "AM" + DateTime.Now.ToString("yyyyMMdd");
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
            string kyes = NumberName + AutoNum + "X";
            
            System.Data.DataTable dt2 = mail.Account_LC3;


                DataRow drw2 = dt2.NewRow();
                drw2["LCCODE"] = lCCODETextBox.Text;
                drw2["LC3CODE"] = kyes;
                dt2.Rows.Add(drw2);
                int T1 = Convert.ToInt16(textBox8.Text);
                if (T1 > 1)
                {
                    for (int i = 0; i <= T1 - 2; i++)
                    {

                        drw2 = dt2.NewRow();
                        drw2["LCCODE"] = lCCODETextBox.Text;
                        drw2["LC3CODE"] = kyes;
                        dt2.Rows.Add(drw2);
                    }
                }
                this.account_LC3BindingSource.EndEdit();

        }

        private void account_LC3DataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (globals.GroupID.ToString().Trim() == "SHI" || globals.GroupID.ToString().Trim() == "EEP" || globals.GroupID.ToString().Trim() == "ACC")
            {
                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                {
                    string sd = account_LC3DataGridView.CurrentRow.Cells["SHIPPINGCODE"].Value.ToString();

                    if (sd == "")
                    {
                        MessageBox.Show("�п�JCargo Receipt No");
                        return;
                    }

                    System.Data.DataTable dt1 = GetOrderDataAPL(sd);
                    if (dt1.Rows.Count > 0)
                    {
                        for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                        {
                            DataRow drw = dt1.Rows[i];
                            System.Diagnostics.Process.Start(drw["LINK"].ToString());
                        }
                        DataGridViewLinkCell cell =

                            (DataGridViewLinkCell)dgv[e.ColumnIndex, e.RowIndex];
                    }
                    else
                    {
                        MessageBox.Show("�S�����");
                    }
                }
            }
            else
            {
                MessageBox.Show("�z�S���[�ݦ��ɮ��v��");
            }
        }


        private System.Data.DataTable GetOrderDataAPL(string SHIPPINGCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT SHIPPINGCODE,[PATH] LINK FROM download3 WHERE SHIPPINGCODE=@SHIPPINGCODE ");
           

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", SHIPPINGCODE));

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

        private void dRAFTDayTextBox_TextChanged(object sender, EventArgs e)
        {

            string H1 = dRAFTDayTextBox.Text;
            int G1 = H1.IndexOf("30");
            int G2 = H1.IndexOf("60");
            int G3 = H1.IndexOf("90");
            int G4 = H1.IndexOf("180");
            if (G1 != -1)
            {

                lCTYPETextBox.Text = "30 days";
            }
            else if (G2 != -1)
            {

                lCTYPETextBox.Text = "60 days";
            }
            else if (G3 != -1)
            {

                lCTYPETextBox.Text = "90 days";
            }
            else if (G4 != -1)
            {

                lCTYPETextBox.Text = "180 days";
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            oCCURTextBox.Text = comboBox2.Text;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            dRAFTTextBox.Text = comboBox3.Text;
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            cOMPANYTextBox.Text = comboBox4.Text;
        }











    }
      
    }



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
namespace ACME
{
    public partial class TTARMAS : ACME.fmBase1
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public TTARMAS()
        {
            InitializeComponent();
        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            sATTARMASTableAdapter.Connection = MyConnection;
            sATT1ARMASTableAdapter.Connection = MyConnection;
            sATT2ARMASTableAdapter.Connection = MyConnection;


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

                    decimal NTDAMOUNT = Convert.ToDecimal(row["NTDAMOUNT"]);
                    UpdateTTUSD(NTDAMOUNT, id, tTCodeTextBox.Text);
                    System.Data.DataTable dt2 = GetTT2(id, tTCodeTextBox.Text);
                    if (dt2.Rows.Count > 0)
                    {
                        string G1 = dt2.Rows[0][0].ToString();
                        string G2 = dt2.Rows[0][1].ToString();
                        UpdateTTPAY(G2,G1, id, tTCodeTextBox.Text);
                    }
                }
                sATT1ARMASTableAdapter.Fill(sa.SATT1ARMAS, MyID);


       

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void UpdateTTUSD(decimal TTUSD, string id, string ttcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update SATT1ARMAS set tttotal=@TTUSD-TotalAmount,LCNTD=@TTUSD,Detail='已提供' where seqno=@id and ttcode=@ttcode");

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
        private void UpdateTTPAY(string GB1,string GB2, string id, string ttcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update SATT1ARMAS set GB1=@GB1,GB2=@GB2 where seqno=@id and ttcode=@ttcode");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@GB1", GB1));
            command.Parameters.Add(new SqlParameter("@GB2", GB2));
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
        private void UpdateBILLNO(string BILLNO, string ID1)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" update SATT2ARMAS set BILLNO=@BILLNO where ID1=@ID1");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@BILLNO", BILLNO));
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

        public static System.Data.DataTable GetTT2(string ID, string TTCode)
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("        SELECT CAST(CAST((100*(ROUND(TOTALAMOUNT/LCNTD,3) )) AS DECIMAL(10,2)) AS VARCHAR)+'%' 付款,CASE WHEN (ROUND(TOTALAMOUNT /LCNTD,3)) < 0.81 THEN '未收訂'  ");
            sb.Append("                             WHEN (ROUND(TOTALAMOUNT /LCNTD,3)) > 0.81 AND  (ROUND(TOTALAMOUNT /LCNTD,3)) <1 THEN '已收訂'  ");
            sb.Append("                             WHEN (ROUND(TOTALAMOUNT /LCNTD,3)) >= 1 THEN '結清'  END PAY  FROM SATT1ARMAS  ");
            sb.Append(" where seqno=@ID and TTCode=@TTCode ");


            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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
        public static System.Data.DataTable GetTT(string TTCode)
        {
            SqlConnection MyConnection = globals.Connection;
            string sql = "SELECT ID,SUM(NTDAMOUNT) NTDAMOUNT FROM SATT2ARMAS where TTCode=@TTCode GROUP BY ID";
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
        private void UpdateTT1(string ttcode)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE  satt1ARMAS SET TTTOTAL=NULL,Detail='' where tTcOde=@ttcode AND SEQNO IN (SELECT distinct seqno FROM SATT1ARMAS where ttcode=@ttcode and seqno not in (select distinct id from satt2ARMAS where ttcode=@ttcode))");

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
        public override void EndEdit()
        {
            Control();
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();
                sa.SATT1ARMAS.RejectChanges();
                sa.SATT2ARMAS.RejectChanges();

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
            tTDateTextBox.ReadOnly = false;
        }
        private void Control()
        {

        }
        public override void AfterEdit()
        {
            tTDateTextBox.ReadOnly = true;
        }
        public override void SetDefaultValue()
        {

            string NumberName = "TA" + DateTime.Now.ToString("yyyy");
            string AutoNum = util.GetAutoNumber(MyConnection, NumberName);

            this.tTCodeTextBox.Text = NumberName + AutoNum;
            tTDateTextBox.Text = DateTime.Now.ToString("yyyyMMdd");

            this.sATTARMASBindingSource.EndEdit();
        }
        public override void SetInit()
        {

            MyBS = sATTARMASBindingSource;
            MyTableName = "SATTARMAS";
            MyIDFieldName = "TTCode";
        }
        public override void FillData()
        {

            try
            {
                sATTARMASTableAdapter.Fill(sa.SATTARMAS, MyID);
                sATT1ARMASTableAdapter.Fill(sa.SATT1ARMAS, MyID);
                sATT2ARMASTableAdapter.Fill(sa.SATT2ARMAS, MyID);

                if (sa.SATT2ARMAS.Rows.Count > 0)
                {
                    for (int i = 0; i <= sa.SATT2ARMAS.Rows.Count - 1; i++)
                    {
                        string BILLNO = sa.SATT2ARMAS.Rows[i]["BILLNO"].ToString();
                        string Docentry = sa.SATT2ARMAS.Rows[i]["Docentry"].ToString();
                        string ID1 = sa.SATT2ARMAS.Rows[i]["ID1"].ToString();
                        
                        if (string.IsNullOrEmpty(BILLNO))
                        {
                            System.Data.DataTable ST2 = GETODLN(Docentry);
                            if (ST2.Rows.Count > 0)
                            {
                                StringBuilder sb = new StringBuilder();
                                for (int f = 0; f <= ST2.Rows.Count - 1; f++)
                                {
                                    DataRow d = ST2.Rows[f];
                                    sb.Append(d["BILLNO"].ToString() + "/");
                                }
                                sb.Remove(sb.Length - 1, 1);
                                UpdateBILLNO(sb.ToString(), ID1);

                            }
                        }


                    }
                    sATT2ARMASTableAdapter.Fill(sa.SATT2ARMAS, MyID);
                }

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

                sATTARMASTableAdapter.Connection.Open();



                Validate();

                sATTARMASBindingSource.EndEdit();
                sATT1ARMASBindingSource.EndEdit();
                sATT2ARMASBindingSource.EndEdit();


                ///注意: 4. 啟動 Transaction

                tx = sATTARMASTableAdapter.Connection.BeginTransaction();



                SqlDataAdapter oWhsAdapter = GetAdapter(sATTARMASTableAdapter);
                oWhsAdapter.UpdateCommand.Transaction = tx;
                oWhsAdapter.InsertCommand.Transaction = tx;
                oWhsAdapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter oWhsAdapter1 = GetAdapter(sATT1ARMASTableAdapter);
                oWhsAdapter1.UpdateCommand.Transaction = tx;
                oWhsAdapter1.InsertCommand.Transaction = tx;
                oWhsAdapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter oWhsAdapter2 = GetAdapter(sATT2ARMASTableAdapter);
                oWhsAdapter2.UpdateCommand.Transaction = tx;
                oWhsAdapter2.InsertCommand.Transaction = tx;
                oWhsAdapter2.DeleteCommand.Transaction = tx;




                sATTARMASTableAdapter.Update(sa.SATTARMAS);
                sATT1ARMASTableAdapter.Update(sa.SATT1ARMAS);
                sATT2ARMASTableAdapter.Update(sa.SATT2ARMAS);


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

                MessageBox.Show(ex.Message, "更新錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                UpdateData = false;
                return UpdateData;
            }
            finally
            {
                this.sATTARMASTableAdapter.Connection.Close();

            }
            return UpdateData;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (sATT1ARMASDataGridView.SelectedRows.Count == 0 )
            {
                MessageBox.Show("請選擇");
                return;
            }
            TTS frm1 = new TTS();
           
            if (frm1.ShowDialog() == DialogResult.OK)
            {
                System.Data.DataTable dt1 = GetAR2(frm1.a);
                System.Data.DataTable dt2 = sa.SATT2ARMAS;

                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();

                    
                    
                    string da = sATT1ARMASDataGridView.SelectedRows[0].Cells["Seqno"].Value.ToString();


                    drw2["ttcode"] = tTCodeTextBox.Text;
                    drw2["id"] = da;
                    drw2["DOCDATE"] = drw["日期"];
                    string G1 = drw["客戶訂單號碼"].ToString();
                    drw2["CardCode"] = drw["客戶ID"];
                    drw2["CardName"] = drw["客戶簡稱"];
                    drw2["SALES"] = drw["業務"];
                    drw2["PROJECT"] = drw["ProjectID"];
                    string FF = drw["訂單總金額"].ToString().Replace(",", "");
                    if (String.IsNullOrEmpty(FF))
                    {
                        FF = "0";
                    }
                    drw2["TAMT"] = FF;
                    //System.Data.DataTable ST = GetCUSTTYPE(G1);
                    //if (ST.Rows.Count > 0)
                    //{
                    //    drw2["SOURCE"] = ST.Rows[0][0].ToString();
                    //}
                    string DOCENTRY = drw["訂單號碼"].ToString();
                    drw2["Docentry"] = DOCENTRY;
                    drw2["NTDAmount"] = drw["訂單金額"];
                    drw2["ShipDate"] = drw["取貨日"];
                    drw2["CUSTCODE"] = drw["外部訂單單號"];
                    drw2["DUETO"] = drw["DUETO"];

                    
                    StringBuilder sb = new StringBuilder();
                    System.Data.DataTable ST2 = GETODLN(DOCENTRY);
                    if (ST2.Rows.Count > 0)
                    {
                        for (int f = 0; f <= ST2.Rows.Count - 1; f++)
                        {
                            DataRow d = ST2.Rows[f];
                            sb.Append(d["BILLNO"].ToString() + "/");
                        }
                        sb.Remove(sb.Length - 1, 1);
                        drw2["BILLNO"] = sb.ToString();
                    }
                    dt2.Rows.Add(drw2);

                }
            }
        }
        public  System.Data.DataTable GetAR2(string DocEntry)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT BILLNO 訂單號碼,BillDate 日期,P.PersonName 業務,T0.SumAmtATax 訂單金額,CustBillNo 客戶訂單號碼,T2.ShortName 客戶簡稱,T0.CustomerID 客戶ID,T0.UserDef1 取貨日,T0.ProjectID, ");
            sb.Append(" REPLACE(SUBSTRING(Remark,CHARINDEX('外部訂單總金額:',Remark),CHARINDEX('5.付款人:',Remark)-CHARINDEX('外部訂單總金額:',Remark)),'外部訂單總金額:','') 訂單總金額");
            sb.Append(" ,REPLACE(SUBSTRING(Remark,CHARINDEX('外部訂單單號:',Remark),CHARINDEX('4.外部訂單總金額:',Remark)-CHARINDEX('外部訂單單號:',Remark)),'外部訂單單號:','') 外部訂單單號,T0.DUETO ");
            sb.Append("    FROM OrdBillMain T0");
            sb.Append("      left join comPerson P ON (T0.Salesman=P.PersonID)");
            sb.Append("       Inner Join comCustomer T2 ON (T0.CustomerID=T2.ID) ");
            sb.Append(" WHERE BILLNO IN  (" + DocEntry + ")");
            
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
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

        //public  System.Data.DataTable GetCUSTTYPE(string CUSTNO)
        //{
        //    SqlConnection MyConnection = globals.Connection;
        //    StringBuilder sb = new StringBuilder();
        //    sb.Append(" SELECT CUSTTYPE FROM GB_POTATO  WHERE ID=@CUSTNO ");

        //    SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
        //    command.CommandType = CommandType.Text;
        //    command.Parameters.Add(new SqlParameter("@CUSTNO", CUSTNO));
        //    SqlDataAdapter da = new SqlDataAdapter(command);
        //    DataSet ds = new DataSet();
        //    try
        //    {
        //        MyConnection.Open();
        //        da.Fill(ds, " inv1 ");
        //    }
        //    finally
        //    {
        //        MyConnection.Close();
        //    }
        //    return ds.Tables[" inv1 "];
        //}
        public System.Data.DataTable GETODLN(string FromNO)
        {
            SqlConnection MyConnection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT BILLNO FROM   ComProdRec A            ");
            sb.Append(" Left Join comWareHouse D On D.WareHouseID=A.WareID");
            sb.Append(" WHERE Flag =500 AND FromNO =@FromNO ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@FromNO", FromNO));
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

        private void sATT1ARMASDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["SeqNo"].Value = util.GetSeqNo(2, sATT1ARMASDataGridView);
        }

        private void sATT1ARMASDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (sATT1ARMASDataGridView.Columns[e.ColumnIndex].Name == "Amount" )
                {

                    decimal Amount = 0;
                    Amount = Convert.ToDecimal(this.sATT1ARMASDataGridView.Rows[e.RowIndex].Cells["Amount"].Value);
                    this.sATT1ARMASDataGridView.Rows[e.RowIndex].Cells["NTD2"].Value = (Amount).ToString("0");

                }

                if (sATT1ARMASDataGridView.Columns[e.ColumnIndex].Name == "Amount" ||
                          sATT1ARMASDataGridView.Columns[e.ColumnIndex].Name == "Fee")
                {

                    string FF = this.sATT1ARMASDataGridView.Rows[e.RowIndex].Cells["Fee"].Value.ToString();
                    if (String.IsNullOrEmpty(FF))
                    {
                        FF = "0";
                    }
                    decimal Amount = 0;
                    decimal Fee = 0;
                    Amount = Convert.ToDecimal(this.sATT1ARMASDataGridView.Rows[e.RowIndex].Cells["Amount"].Value);
                    Fee = Convert.ToDecimal(FF);

                    sATT1ARMASDataGridView.Rows[e.RowIndex].Cells["TotalAmount"].Value = (Amount + Fee).ToString();

                }
                if (sATT1ARMASDataGridView.Columns[e.ColumnIndex].Name == "TTCheck" ||
                   sATT1ARMASDataGridView.Columns[e.ColumnIndex].Name == "BankCheck" ||
                   sATT1ARMASDataGridView.Columns[e.ColumnIndex].Name == "ALCheck")
                {
                    string TTCheck = sATT1ARMASDataGridView.Rows[e.RowIndex].Cells["TTCheck"].Value.ToString();
                    string BankCheck = sATT1ARMASDataGridView.Rows[e.RowIndex].Cells["BankCheck"].Value.ToString();
                    DateTime P1 = DateTime.ParseExact(TTCheck, "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);
                    string DATETIME = "";
                    if (BankCheck == "True")
                    {
                        DATETIME = P1.AddDays(7).ToString("yyyyMMdd");
                    }
                    else
                    {
                        DATETIME = P1.AddDays(3).ToString("yyyyMMdd");
                    }

                    sATT1ARMASDataGridView.Rows[e.RowIndex].Cells["ALCheck"].Value = DATETIME;
                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
        }

        private void sATT1ARMASDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void sATT2ARMASDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }
    }
}

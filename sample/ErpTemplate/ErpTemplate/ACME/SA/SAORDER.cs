using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.Collections;
using System.IO;

namespace ACME
{
    public partial class SAORDER : Form
    {
        string FA = "acmesql02";
       
        public SAORDER()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if (comboBox2.SelectedValue.ToString() == "Please-Select")
            //{
            //    MessageBox.Show("請選擇業助");
            //    return;
            //}
            System.Data.DataTable G2 = GetORDR();
            dataGridView1.DataSource = G2;
            dataGridView1.Columns[0].ReadOnly = true;
            dataGridView1.Columns[1].ReadOnly = true;
            dataGridView1.Columns[2].ReadOnly = true;
            dataGridView1.Columns[3].ReadOnly = true;
            dataGridView1.Columns[4].ReadOnly = true;
            dataGridView1.Columns[5].ReadOnly = true;
            dataGridView1.Columns[6].ReadOnly = true;
        }

        public  System.Data.DataTable GetORDR()
        {
            SqlConnection MyConnection = globals.shipConnection;
            string OHEM = fmLogin.LoginID.ToString().ToUpper();
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT T0.DOCENTRY SO,T1.LINENUM LINENUM,T0.CARDNAME Client,T0.NUMATCARD PO,T1.Dscription Model,CAST(T1.QUANTITY AS INT) Qty,T0.DocCur +' ' +CAST(CAST(ROUND(T1.Price,4) AS DECIMAL(14,4)) AS VARCHAR) Price    ");
            sb.Append(" ,Convert(varchar(10), T1.U_ACME_SHIPDAY,112) 離倉日期    ");
            sb.Append(" ,Convert(varchar(10), T1.U_ACME_WORK,112) 排程日期    ");
            //u_shipday
            sb.Append(" ,T1.U_PAY Payment,T1.U_SHIPSTATUS 貨況    ");

            if (OHEM == "ESTHERYEH"||OHEM == "CLAIRECHEN")
            {
                sb.Append(" ,T1.u_shipday 押出貨日    ");
            }
            sb.Append(" ,T1.U_ACME_Dscription 'SA備註(SI)',T1.U_MEMO 注意事項,T2.DOCDATE 時間,T2.REMARK 事項  FROM ORDR T0    ");
            sb.Append(" LEFT JOIN RDR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)    ");
            sb.Append(" LEFT JOIN ACMESQLSP.DBO.SA_ORDER T2 ON (T1.DocEntry=T2.DOCENTRY AND T1.LINENUM =T2.LINENUM)  ");
            sb.Append(" left JOIN OHEM T3 ON T0.OwnerCode = T3.empID   ");
            sb.Append(" iNNER JOIN OWHS T4 ON T4.whsCode = T1.whscode  ");
    
            sb.Append(" WHERE T1.LineStatus ='O'    ");

            if (OHEM == "LLEYTONCHEN")
            {
                sb.Append(" and T3.HOMETEL='ESTHERYEH' ");
            }
            else
            {
                sb.Append(" and T3.HOMETEL=@lastName ");
            }
         //   sb.Append(" and T3.[lastName]+T3.[firstName]=@lastName ");

            if (textBox1.Text != "" && textBox3.Text != "")
            {
                sb.Append("and  Convert(varchar(8),T1.U_ACME_WORK,112)  BETWEEN '" + textBox1.Text.ToString() + "' AND  '" + textBox3.Text.ToString() + "'    ");
            }
            if (textBox2.Text != "")
            {
                sb.Append(" and  t0.[cardname] like N'%" + textBox2.Text.ToString() + "%'  ");
            }

            if (textBox4.Text != "")
            {
                sb.Append(" and  T1.Dscription  like N'%" + textBox4.Text.ToString() + "%'  ");
            }

            if (textBox5.Text != "")
            {
                sb.Append(" and  T1.U_ACME_Dscription  like N'%" + textBox5.Text.ToString() + "%'  ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@lastName", fmLogin.LoginID.ToString()));
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
        private System.Data.DataTable GetSO(int DOCENTRY, int LINENUM)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT LINENUM FROM SA_ORDER WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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
        private System.Data.DataTable GetSO2(int DOCENTRY, int LINENUM,string DOCDATE)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT Convert(varchar(10),dbo.fun_SHIPDATE(U_ACME_WORKDAY,@DOCDATE),112)   FROM RDR1 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
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
        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {

                DataGridViewRow row;

                row = dataGridView1.Rows[i];
                int SO = Convert.ToInt32(row.Cells["SO"].Value);
                int LINENUM = Convert.ToInt16(row.Cells["LINENUM"].Value);
         //       string PO = row.Cells["PO"].Value.ToString();
           //     double Qty = Convert.ToDouble(row.Cells["Qty"].Value);
          //      double Price = Convert.ToDouble(row.Cells["Price"].Value);


          
                string Deliverydate = row.Cells["排程日期"].Value.ToString().Substring(0, 4) + "/" + row.Cells["排程日期"].Value.ToString().Substring(4, 2) + "/" + row.Cells["排程日期"].Value.ToString().Substring(6, 2);
                string LEAVEDAY = row.Cells["離倉日期"].Value.ToString().Substring(0, 4) + "/" + row.Cells["離倉日期"].Value.ToString().Substring(4, 2) + "/" + row.Cells["離倉日期"].Value.ToString().Substring(6, 2);

                DateTime LDATE = Convert.ToDateTime(LEAVEDAY);
                DateTime DDATE = Convert.ToDateTime(Deliverydate);
                string Payment = row.Cells["Payment"].Value.ToString();
                string Status = row.Cells["貨況"].Value.ToString();
                string SI = row.Cells["SA備註(SI)"].Value.ToString();
                string Remarks = row.Cells["注意事項"].Value.ToString();
                string 時間 = row.Cells["時間"].Value.ToString();
                string 事項 = row.Cells["事項"].Value.ToString();
                string OHEM = fmLogin.LoginID.ToString().ToUpper();
                if (OHEM == "ESTHERYEH" || OHEM == "CLAIRECHEN")
                {
                    string 押出貨日 = row.Cells["押出貨日"].Value.ToString();

                    UPDATEORDER3(押出貨日, SO, LINENUM);
                
                }

                UPDATEORDER2( LDATE, DDATE, Payment, Status, SI, Remarks,SO,LINENUM);

                System.Data.DataTable GO = GetSO(SO, LINENUM);

                if (GO.Rows.Count > 0)
                {

                    UPDATEORDER(時間, 事項, SO, LINENUM);
                }
                else
                {

                    AddORDER(時間, 事項, SO, LINENUM);
                }

            }

            //int F = 0;
            //SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

            //oCompany = new SAPbobsCOM.Company();

            //oCompany.Server = "acmesap";
            //oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            //oCompany.UseTrusted = false;
            //oCompany.DbUserName = "sapdbo";
            //oCompany.DbPassword = "@rmas";
            //oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;


            //oCompany.CompanyDB = FA;
            //oCompany.UserName = "S03";
            //oCompany.Password = "0108";
            //int result = oCompany.Connect();
            //if (result == 0)
            //{
            //    SAPbobsCOM.Documents oORDER = null;
            //    oORDER = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
            //    SAPbobsCOM.Document_Lines ORDERLINE = null;




            //    for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            //    {

            //        DataGridViewRow row;

            //        row = dataGridView1.Rows[i];
            //        int SO = Convert.ToInt32(row.Cells["SO"].Value);
            //        int LINENUM = Convert.ToInt16(row.Cells["LINENUM"].Value);
            //        string PO = row.Cells["PO"].Value.ToString();
            //        double Qty = Convert.ToDouble(row.Cells["Qty"].Value);
            //        double Price = Convert.ToDouble(row.Cells["Price"].Value);


            //        string Expected = row.Cells["訂單交期"].Value.ToString().Substring(0, 4) + "/" + row.Cells["訂單交期"].Value.ToString().Substring(4, 2) + "/" + row.Cells["訂單交期"].Value.ToString().Substring(6, 2);
            //        string Deliverydate = row.Cells["排程日期"].Value.ToString().Substring(0, 4) + "/" + row.Cells["排程日期"].Value.ToString().Substring(4, 2) + "/" + row.Cells["排程日期"].Value.ToString().Substring(6, 2);
            //        string LEAVEDAY = row.Cells["離倉日期"].Value.ToString().Substring(0, 4) + "/" + row.Cells["離倉日期"].Value.ToString().Substring(4, 2) + "/" + row.Cells["離倉日期"].Value.ToString().Substring(6, 2);
            //        DateTime SHIPDATE = Convert.ToDateTime(Deliverydate);
            //        DateTime LDATE = Convert.ToDateTime(LEAVEDAY);
            //        DateTime DDATE = Convert.ToDateTime(Deliverydate);
            //        string Payment = row.Cells["Payment"].Value.ToString();
            //        string Status = row.Cells["Status"].Value.ToString();
            //        string SI = row.Cells["SI"].Value.ToString();
            //        string Remarks = row.Cells["Remarks"].Value.ToString();
            //        string 時間 = row.Cells["時間"].Value.ToString();
            //        string 事項 = row.Cells["事項"].Value.ToString();
            //        if (oORDER.GetByKey(SO))
            //        {
            //            ORDERLINE = oORDER.Lines;
            //            ORDERLINE.Add();

            //            ORDERLINE.SetCurrentLine(LINENUM);
            //            ORDERLINE.Quantity = Qty;
            //            oORDER.NumAtCard = PO;
            //            ORDERLINE.Price = Price;

            //            ORDERLINE.ShipDate = SHIPDATE;
            //            ORDERLINE.UserFields.Fields.Item("U_ACME_WORK").Value = DDATE;
            //            ORDERLINE.UserFields.Fields.Item("U_ACME_SHIPDAY").Value = LDATE;
            //            ORDERLINE.UserFields.Fields.Item("U_PAY").Value = Payment;
            //            ORDERLINE.UserFields.Fields.Item("U_SHIPSTATUS").Value = Status;
            //            ORDERLINE.UserFields.Fields.Item("U_ACME_Dscription").Value = SI;
            //            ORDERLINE.UserFields.Fields.Item("U_MEMO").Value = Remarks;
            //            int res = oORDER.Update();
            //            if (res != 0)
            //            {
            //                F = 1;
            //                MessageBox.Show("修改錯誤 " + oCompany.GetLastErrorDescription());
            //                return;
            //            }
            //            else
            //            {
            //                System.Data.DataTable GO = GetSO(SO, LINENUM);

            //                if (GO.Rows.Count > 0)
            //                {

            //                    UPDATEORDER(時間, 事項, SO, LINENUM);
            //                }
            //                else
            //                {

            //                    AddORDER(時間, 事項, SO, LINENUM);
            //                }


            //            }
            //        }
            //    }

            //    if (F == 0)
            //    {
            //        MessageBox.Show("存檔成功");
            //    }
            //}






            //else
            //{
            //    MessageBox.Show(oCompany.GetLastErrorDescription());

            //}
            MessageBox.Show("存檔成功");


        }


        public void AddORDER(string DOCDATE, string REMARK, int DOCENTRY, int LINENUM)
        {
            SqlConnection connection = globals.Connection;
            SqlCommand command = new SqlCommand("Insert into SA_ORDER(DOCENTRY,LINENUM,DOCDATE,REMARK) values(@DOCENTRY,@LINENUM,@DOCDATE,@REMARK)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));

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

        private void UPDATEORDER(string DOCDATE, string REMARK, int DOCENTRY, int LINENUM)
        {


            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE SA_ORDER SET DOCDATE =@DOCDATE,REMARK=@REMARK WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);

            command.Parameters.Add(new SqlParameter("@DOCDATE", DOCDATE));
            command.Parameters.Add(new SqlParameter("@REMARK", REMARK));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));


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
        private void UPDATEORDER2(DateTime U_ACME_SHIPDAY, DateTime U_ACME_WORK, string U_PAY, string U_SHIPSTATUS, string U_ACME_Dscription, string U_MEMO, int DOCENTRY, int LINENUM)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE RDR1 SET U_ACME_SHIPDAY=@U_ACME_SHIPDAY,U_ACME_WORK=@U_ACME_WORK,U_PAY=@U_PAY,U_SHIPSTATUS=@U_SHIPSTATUS,U_ACME_Dscription=@U_ACME_Dscription,U_MEMO=@U_MEMO WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            command.Parameters.Add(new SqlParameter("@U_ACME_SHIPDAY", U_ACME_SHIPDAY));
            command.Parameters.Add(new SqlParameter("@U_ACME_WORK", U_ACME_WORK));
            command.Parameters.Add(new SqlParameter("@U_PAY", U_PAY));

            command.Parameters.Add(new SqlParameter("@U_SHIPSTATUS", U_SHIPSTATUS));
            command.Parameters.Add(new SqlParameter("@U_ACME_Dscription", U_ACME_Dscription));
            command.Parameters.Add(new SqlParameter("@U_MEMO", U_MEMO));

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));


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
        private void UPDATEORDER3(string u_shipday, int DOCENTRY, int LINENUM)
        {


            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE RDR1 SET u_shipday=@u_shipday  WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);


            command.Parameters.Add(new SqlParameter("@u_shipday", u_shipday));
          
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));


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
        private void SAORDER_Load(object sender, EventArgs e)
        {
       //     UtilSimple.SetLookupBinding(comboBox2, GetMenu.GetOhem(), "DataValue", "DataValue");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView1);
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                if (dataGridView1.Columns[e.ColumnIndex].Name == "離倉日期")
                {
                    int SO = Convert.ToInt32(this.dataGridView1.Rows[e.RowIndex].Cells["SO"].Value);
                    int LINENUM = Convert.ToInt32(this.dataGridView1.Rows[e.RowIndex].Cells["LINENUM"].Value);
                    string 離倉日期 = this.dataGridView1.Rows[e.RowIndex].Cells["離倉日期"].Value.ToString().Substring(0, 4) + "/" + this.dataGridView1.Rows[e.RowIndex].Cells["離倉日期"].Value.ToString().Substring(4, 2) + "/" + this.dataGridView1.Rows[e.RowIndex].Cells["離倉日期"].Value.ToString().Substring(6, 2);
                    System.Data.DataTable F1 = GetSO2(SO, LINENUM, 離倉日期);
                    this.dataGridView1.Rows[e.RowIndex].Cells["排程日期"].Value = (F1.Rows[0][0]).ToString();
                    //decimal iQuantity = 0;
                    //decimal iUnitPrice = 0;

                    //iQuantity = Convert.ToInt32(this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["Quantity"].Value);
                    //iUnitPrice = Convert.ToDecimal(this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["ItemPrice"].Value);
                    //this.shipping_ItemDataGridView.Rows[e.RowIndex].Cells["ItemAmount"].Value = (iQuantity * iUnitPrice).ToString();

                }

            }
            catch
            {

            }
        }

    }
}

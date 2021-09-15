using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
namespace ACME
{
    public partial class GROUPM : ACME.fmBase1
    {
        public GROUPM()
        {
            InitializeComponent();
        }
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            gROUPMTableAdapter.Connection = MyConnection;
            gROUPDTableAdapter.Connection = MyConnection;
            gROUPD1TableAdapter.Connection = MyConnection;
        }
        private void WW()
        {
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();

                uSERS.GROUPM.RejectChanges();
                uSERS.GROUPD.RejectChanges();
                uSERS.GROUPD1.RejectChanges();
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
        public override void AfterAddNew()
        {
            WW();
        }
        public override void SetInit()
        {

            MyBS = gROUPMBindingSource;
            MyTableName = "GROUPM";
            MyIDFieldName = "ShippingCode";

        }
        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "BO" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.shippingCodeTextBox.Text = kyes;

            this.gROUPMBindingSource.EndEdit();
            kyes = null;
        }
        public override void FillData()
        {
            try
            {

                gROUPMTableAdapter.Fill(uSERS.GROUPM, MyID);
                gROUPDTableAdapter.Fill(uSERS.GROUPD, MyID);
                gROUPD1TableAdapter.Fill(uSERS.GROUPD1, MyID);
                System.Data.DataTable ggf=GetWAR6();
                UPDATECHECK(ggf.Rows[0][0].ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        public override bool UpdateData()
        {
            bool UpdateData;
            SqlTransaction tx = null;
            try
            {


                Validate();



                gROUPMTableAdapter.Connection.Open();


                gROUPMBindingSource.EndEdit();
                gROUPDBindingSource.EndEdit();
                gROUPD1BindingSource.EndEdit();

                tx = gROUPMTableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(gROUPMTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter1 = util.GetAdapter(gROUPDTableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter2 = util.GetAdapter(gROUPD1TableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;

                gROUPMTableAdapter.Update(uSERS.GROUPM);
                uSERS.GROUPM.AcceptChanges();

                gROUPDTableAdapter.Update(uSERS.GROUPD);
                uSERS.GROUPD.AcceptChanges();

                gROUPD1TableAdapter.Update(uSERS.GROUPD1);
                uSERS.GROUPD1.AcceptChanges();

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
                this.gROUPMTableAdapter.Connection.Close();

            }
            return UpdateData;
        }

        private void gROUPDDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = gROUPDDataGridView.Rows.Count;

            e.Row.Cells["LINENUM"].Value = iRecs.ToString();
        }

        private void gROUPD1DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            int iRecs;

            iRecs = gROUPD1DataGridView.Rows.Count;

            e.Row.Cells["LINENUM1"].Value = iRecs.ToString();
        }

        private void gROUPDDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void gROUPD1DataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

        }

        private void GROUPM_Load(object sender, EventArgs e)
        {
          //  UPDATECHECK 
        }
        private System.Data.DataTable GetWAR6()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @name varchar(500) ");
            sb.Append(" select @name =SUBSTRING(COALESCE(@name + '/',''),0,500) + USERS ");
            sb.Append(" FROM ( SELECT   CAST(ROW_NUMBER() OVER (ORDER BY USERS )  AS VARCHAR)+'.'+USERS+' '+CAST(T1.AMT AS VARCHAR)+'K'  USERS FROM GROUPM T0 ");
            sb.Append(" LEFT JOIN GROUPD1 T1 ON (T0.ShippingCode =T1.ShippingCode)");
            sb.Append(" WHERE T0.ShippingCode =@ShippingCode) PS");
            sb.Append(" SELECT @name MEMBER");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
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
        public void UPDATECHECK(string MEMBER)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" UPDATE GROUPM SET MEMBER=@MEMBER WHERE SHIPPINGCODE=@SHIPPINGCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@MEMBER", MEMBER));
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

        private void button2_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(gROUPD1DataGridView);
        }
    }
}

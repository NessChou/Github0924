using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.VisualBasic;
namespace ACME
{
    public partial class fmDsmenu : Form
    {
        string str16 = "Data Source=10.10.1.40;Initial Catalog=CHICOMP16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        string EEP = "Data Source=acmesap;Initial Catalog=acmesqlEEP;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string DRS = "Data Source=acmesap;Initial Catalog=acmesql05;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        string HR104 = "Data Source=10.10.1.45;Initial Catalog=89206602;Persist Security Info=True;User ID=ehrview;Password=viewehr";
        string FA = "acmesql98";

        public fmDsmenu()
        {
            InitializeComponent();
        }

        private void mENUTABLEBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.mENUTABLEBindingSource.EndEdit();
            this.mENUTABLETableAdapter.Update(this.uSERS.MENUTABLE);

            MessageBox.Show("儲存成功");

        }

        private void fmDsmenu_Load(object sender, EventArgs e)
        {

            // TODO: 這行程式碼會將資料載入 'uSERS.RPA_PackingD' 資料表。您可以視需要進行移動或移除。


            // TODO: 這行程式碼會將資料載入 'uSERS.RMA_PARAMS' 資料表。您可以視需要進行移動或移除。
            this.rMA_PARAMSTableAdapter.Fill(this.uSERS.RMA_PARAMS);
            // TODO: 這行程式碼會將資料載入 'uSERS.PARAMS' 資料表。您可以視需要進行移動或移除。
            this.pARAMSTableAdapter.Fill(this.uSERS.PARAMS);
            // TODO: 這行程式碼會將資料載入 'uSERS.right_group' 資料表。您可以視需要進行移動或移除。
            this.right_groupTableAdapter.Fill(this.uSERS.right_group);
            this.uSERSSHIPTableAdapter.Fill(this.uSERS.USERSSHIP);
            this.aCME_AUTOPROTableAdapter.Fill(this.uSERS.ACME_AUTOPRO);
            this.aCME_ARES_MAILTableAdapter.Fill(this.mail.ACME_ARES_MAIL);

            comboBox8.Text = "SAP";
            UtilSimple.SetLookupBinding(comboBox1, Getpdn(), "DataText", "DataValue");
            UtilSimple.SetLookupBinding(comboBox2, Getpdn(), "DataText", "DataValue");

            this.uSERMENUSTableAdapter.Fill(this.uSERS.USERMENUS);
            this.uSERSSHIPTableAdapter.Fill(this.uSERS.USERSSHIP);
            this.uSERSTableAdapter.Fill(this.uSERS._USERS);
            this.rightTableAdapter.Fill(this.uSERS.Right);
            this.mENUTABLETableAdapter.Fill(this.uSERS.MENUTABLE);

            dataGridView1.DataSource = GetSCE();
            yEARToolStripTextBox.Text = GetMenu.DayYEAR();
            UtilSimple.SetLookupBinding(comboBox6, Getpdn1(), "DataText", "DataValue");
            UtilSimple.SetLookupBinding(comboBox3, GetOslp(), "DataText", "DataValue");
            UtilSimple.SetLookupBinding(comboBox4, GetOslp(), "DataText", "DataValue");
            comboBox7.DataSource = Getemployee();
            comboBox7.DisplayMember = "PARAM_NO";
            comboBox7.ValueMember = "PARAM_NO";

            comboBox5.DataSource = Getemployee();
            comboBox5.DisplayMember = "PARAM_NO";
            comboBox5.ValueMember = "PARAM_NO";

        }
        public static System.Data.DataTable Getpdn1()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = " SELECT DOCENTRY as datavalue, FormName as datatext FROM FORMID  union all  SELECT 0   as DataValue,'Please-Select'   as Datatext  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM'  ORDER BY DOCENTRY   ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "FORMID");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["FORMID"];
        }
        System.Data.DataTable GetOslp()
        {

            SqlConnection con = globals.shipConnection;
            string sql = "  select userid as datavalue,user_code as datatext from ousr  union all  SELECT 0   as DataValue,'Please-Select'   as Datatext  FROM ACMESQLSP.DBO.PARAMS WHERE PARAM_KIND='OHEM' order by userid  ";


            SqlDataAdapter da = new SqlDataAdapter(sql, con);
            DataSet ds = new DataSet();
            try
            {
                con.Open();
                da.Fill(ds, "ousr");
            }
            finally
            {
                con.Close();
            }
            return ds.Tables["ousr"];
        }
        public static System.Data.DataTable Getemployee()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from RMA_PARAMS WHERE PARAM_KIND='DB' ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
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
        public static System.Data.DataTable Getpdn()
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = " SELECT USERID as datavalue,USERID as datatext FROM USERS union all  SELECT TOP 1'0'   as DataValue,'Please-Select'   as datatext  FROM USERS ORDER BY DataValue  ";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "USERS");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["USERS"];
        }
        private System.Data.DataTable GetSCE()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" select distinct t0.hostname 電腦,t1.[name] 登入者,login_time 登入時間,last_batch 最後動作時間 from master..sysprocesses t0");
            sb.Append(" left join acmesqlsp.dbo.hostname t1 on(t0.hostname=t1.hostname COLLATE Chinese_Taiwan_Stroke_CI_AS)");
            sb.Append(" where program_name='SAP Business One' and dbid=13 ");
            sb.Append(" order by t1.[name] ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
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
            sb.Append("SELECT date_ID ID,DATE_TIME ITIME FROM Y_2004 ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
        private System.Data.DataTable GetCUVV()
        {
            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT [VALUE],LINENUM FROM CUVV WHERE INDEXID=@INDEXID ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@INDEXID", textBox8.Text));
            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rct2");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

    
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.rightBindingSource.EndEdit();
            this.rightTableAdapter.Update(this.uSERS.Right);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text != "Please-Select" && comboBox2.Text != "Please-Select")
            {
                UpdateMasterSQL2();
            }
        }

        private void UpdateMasterSQL2()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @From nvarchar(20) ");
            sb.Append(" Declare @To nvarchar(20) ");

            sb.Append(" set @From = '" + comboBox1.SelectedValue.ToString() + "' ");
            sb.Append(" set @To = '" + comboBox2.SelectedValue.ToString() + "' ");
            sb.Append(" begin");
            sb.Append(" Select * into USERMENUS2  From USERMENUS ");
            sb.Append(" DELETE USERMENUS WHERE USERID=@To ");
            sb.Append(" DELETE USERMENUS2 WHERE USERID=@To ");
            sb.Append(" UPDATE USERMENUS2 SET USERID=@To WHERE USERID=@From ");
            sb.Append(" Insert Into  USERMENUS  Select * From USERMENUS2 WHERE USERID=@To ");
            sb.Append(" Drop Table USERMENUS2 ");
            sb.Append(" end ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            try
            {

                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                    MessageBox.Show("更新成功");
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
            deleteshipping();
            MessageBox.Show("刪除成功");
           
            
            textBox1.Text = "";

        }
        public void deleteshipping()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" delete shipping_main where shippingcode=@aa");
            sb.Append(" delete Shipping_Item where shippingcode=@aa");
            sb.Append(" delete InvoiceM where shippingcode=@aa");
            sb.Append(" delete InvoiceD where shippingcode=@aa");
            sb.Append(" delete PackingListM where shippingcode=@aa");
            sb.Append(" delete PackingListD where shippingcode=@aa");
            sb.Append(" delete LADINGM where shippingcode=@aa");
            sb.Append(" delete LADINGD where shippingcode=@aa");
            sb.Append(" delete LcInstro where shippingcode=@aa");
            sb.Append(" delete LcInstro1 where shippingcode=@aa");
            sb.Append(" delete WH_main where shippingcode=@aa");
            sb.Append(" delete WH_Item where shippingcode=@aa");
            sb.Append(" delete WH_Item2 where shippingcode=@aa");
            sb.Append(" delete WH_Item3 where shippingcode=@aa");
            sb.Append(" delete WH_Item4 where shippingcode=@aa");
            sb.Append(" delete download where shippingcode=@aa");
            sb.Append(" delete WH_Car2 where shippingcode=@aa");
            sb.Append(" delete WH_PACK2 where shippingcode=@aa");
            sb.Append(" delete WH_PACK3 where shippingcode=@aa");
            sb.Append(" delete cFS where shippingcode=@aa");
            sb.Append(" delete MARK where shippingcode=@aa");
            sb.Append(" DELETE Rma_mainSZ WHERE SHIPPINGCODE=@aa");
            sb.Append(" DELETE Rma_PackingListDSZ WHERE SHIPPINGCODE=@aa");
            sb.Append(" DELETE Rma_InvoiceDSZ WHERE SHIPPINGCODE=@aa");
            sb.Append(" DELETE Rma_main WHERE SHIPPINGCODE=@aa");
            sb.Append(" DELETE Rma_PackingListD WHERE SHIPPINGCODE=@aa");
            sb.Append(" DELETE Rma_InvoiceD WHERE SHIPPINGCODE=@aa");
            sb.Append(" DELETE rMA_LADINGD WHERE SHIPPINGCODE=@aa");

            sb.Append(" delete AcmeSqlSPCHOICE.DBO.WH_main where shippingcode=@aa ");
            sb.Append(" delete AcmeSqlSPCHOICE.DBO.WH_Item where shippingcode=@aa ");
            sb.Append(" delete AcmeSqlSPCHOICE.DBO.WH_Item2 where shippingcode=@aa ");
            sb.Append(" delete AcmeSqlSPCHOICE.DBO.WH_Item3 where shippingcode=@aa ");
            sb.Append(" delete AcmeSqlSPCHOICE.DBO.WH_Item4 where shippingcode=@aa ");
            sb.Append(" delete AcmeSqlSPCHOICE.DBO.WH_Car2 where shippingcode=@aa ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@aa", textBox1.Text));

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
        public void UPDATEOTIM()
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE POR1 SET Dscription =@Dscription WHERE ITEMCODE=@ITEMCODE");
            sb.Append(" UPDATE PDN1 SET Dscription =@Dscription  WHERE ITEMCODE=@ITEMCODE");
            sb.Append(" UPDATE PCH1 SET Dscription =@Dscription WHERE ITEMCODE=@ITEMCODE");
            sb.Append(" UPDATE RDR1 SET Dscription =@Dscription  WHERE ITEMCODE=@ITEMCODE");
            sb.Append(" UPDATE DLN1 SET Dscription =@Dscription  WHERE ITEMCODE=@ITEMCODE");
            sb.Append(" UPDATE INV1 SET Dscription =@Dscription  WHERE ITEMCODE=@ITEMCODE");
            sb.Append(" UPDATE WTR1 SET Dscription =@Dscription  WHERE ITEMCODE=@ITEMCODE");
            sb.Append(" UPDATE IGE1 SET Dscription =@Dscription  WHERE ITEMCODE=@ITEMCODE");
            sb.Append(" UPDATE IGN1 SET Dscription =@Dscription  WHERE ITEMCODE=@ITEMCODE");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", textBox11.Text));
            command.Parameters.Add(new SqlParameter("@Dscription", textBox12.Text));
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

        public void UPCHI16()
        {

            SqlConnection connection = new SqlConnection(str16);
            StringBuilder sb = new StringBuilder();
            sb.Append("UPDATE ComSubject SET IsUseProject =1 WHERE SubjectID = @SubjectID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SubjectID", textBox16.Text));

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
        public void UPWH_FEE()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE WH_FEE SET ENA ='1' WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", textBox14.Text));

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
            int ff = Convert.ToInt32(textBox2.Text);
            AddAUOGD6(ff);
            DateTime dt = new DateTime(ff, 1, 1);
            for (int i = 0; i <= 365; i++)
            {
                string f = "";


                if ((dt.AddDays(i).DayOfWeek.ToString() == "Saturday") || (dt.AddDays(i).DayOfWeek.ToString() == "Sunday"))
                {
                    f = "1";
                }
                else
                {
                    f = "0";
                }
                AddAUOGD5(dt.AddDays(i), f);
            }
            MessageBox.Show("新增成功");
        }

        public void AddAUOGD5(DateTime Date_Time, string IsRestDay)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into Y_2004(Date_Time,IsRestDay) values(@Date_Time,@IsRestDay)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Date_Time", Date_Time));
            command.Parameters.Add(new SqlParameter("@IsRestDay", IsRestDay));
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
        public void UPG5(string ID, int D1, int D2, int D3)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE Y_2004 SET D1=@D1,D2=@D2,D3=@D3  WHERE  date_ID=@ID", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@D1", D1));
            command.Parameters.Add(new SqlParameter("@D2", D2));
            command.Parameters.Add(new SqlParameter("@D3", D3));
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
        public void AddAUOGD5GB(string GBDATE, string STARTDATE, string ENDDATE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into GB_DATE(GBDATE,STARTDATE,ENDDATE) values(@GBDATE,@STARTDATE,@ENDDATE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@GBDATE", GBDATE));
            command.Parameters.Add(new SqlParameter("@STARTDATE", STARTDATE));
            command.Parameters.Add(new SqlParameter("@ENDDATE", ENDDATE));
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
        public void ADDGBWEEK(string SWEEK, string EWEEK, int WNUM)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into GB_FWEEK(SWEEK,EWEEK,WNUM) values(@SWEEK,@EWEEK,@WNUM)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SWEEK", SWEEK));
            command.Parameters.Add(new SqlParameter("@EWEEK", EWEEK));
            command.Parameters.Add(new SqlParameter("@WNUM", WNUM));
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
        public void AddAUOGD6(int Date_Time)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("delete Y_2004 where year(Date_Time)=@Date_Time", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Date_Time", Date_Time));
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
        public void AddAUOGD6GB()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("delete GB_DATE where SUBSTRING(GBDATE,1,4)=@Date_Time", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@Date_Time", textBox9.Text));
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
                ViewBatchPayment2();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ViewBatchPayment2()
        {
            SqlConnection MyConnection = null;
            if (comboBox8.Text =="輔助")
            {
                MyConnection = globals.Connection;

    
            }
            else   if (comboBox8.Text =="SAP")
            {
                MyConnection = globals.shipConnection;

            }
            else if (comboBox8.Text == "EEP")
            {
                MyConnection = new SqlConnection(EEP);

            }
            else if (comboBox8.Text == "104")
            {
                MyConnection = new SqlConnection(HR104);

            }
            else if (comboBox8.Text == "DRS")
            {
                MyConnection = new SqlConnection(DRS);

            }
            StringBuilder sb = new StringBuilder();
            string aa = textBox3.Text;
            SqlCommand command = new SqlCommand(aa, MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPOR");
            }
            finally
            {
                MyConnection.Close();
            }


            dataGridView2.DataSource = ds.Tables[0];

        }


        private void button4_Click(object sender, EventArgs e)
        {
            ExcelReport.GridViewToExcel(dataGridView2);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            UpdateMasterSQL();
            MessageBox.Show("資料已更新");
        }
        private void UpdateMasterSQL()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Drop Table ACMESQLSP.DBO.CPRF ");
            sb.Append(" Select * Into ACMESQLSP.DBO.CPRF From " + comboBox7.Text + ".DBO.CPRF");
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

        private void UpdateMasterSQL22()
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" Declare @FormID nvarchar(20)");
            sb.Append(" Declare @From nvarchar(20)");
            sb.Append(" Declare @To nvarchar(20)");
            sb.Append(" set @FormID = " + comboBox6.SelectedValue.ToString() + "");
            sb.Append(" set @From = " + comboBox3.SelectedValue.ToString() + "");
            sb.Append(" set @To = " + comboBox4.SelectedValue.ToString() + "");
            sb.Append(" if exists(Select 1 from ACMESQL98.DBO.OUSR where UserID=@To)");
            sb.Append(" begin");
            sb.Append(" Delete From  " + comboBox5.SelectedValue.ToString() + ".DBO.CPRF Where (FormID=@FormID Or FormID='-'+@FormID Or @FormID=0) And UserSign=@To");
            sb.Append(" Insert Into  ACMESQLSP.DBO.CPRF2 Select * From ACMESQLSP.DBO.CPRF where (FormID=@FormID Or FormID='-'+@FormID Or @FormID=0) And UserSign=@From");
            sb.Append(" update ACMESQLSP.DBO.CPRF2 set UserSign=@to");
            sb.Append(" Insert Into  " + comboBox5.SelectedValue.ToString() + ".DBO.CPRF Select * From ACMESQLSP.DBO.CPRF2");
            sb.Append(" truncate table ACMESQLSP.DBO.CPRF2 ");
            sb.Append(" end");
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

        private void button6_Click(object sender, EventArgs e)
        {
            if (comboBox3.SelectedValue.ToString() == "0" || comboBox4.SelectedValue.ToString() == "0" || comboBox6.SelectedValue.ToString() == "0")
            {
                MessageBox.Show("請選擇資料");
            }
            else
            {
                UpdateMasterSQL22();
                MessageBox.Show("資料已更新");
            }

        }

        private void button9_Click(object sender, EventArgs e)
        {
            textBox5.Text = UtilSimple.Encrypt(textBox6.Text, "1234");
        }

        private void button8_Click(object sender, EventArgs e)
        {
            textBox4.Text = UtilSimple.Decrypt(textBox5.Text, "1234");
        }

        private void 儲存SToolStripButton_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.rightBindingSource.EndEdit();
            this.rightTableAdapter.Update(this.uSERS.Right);

            this.right_groupBindingSource.EndEdit(); 
            this.right_groupTableAdapter.Update(this.uSERS.right_group);
            MessageBox.Show("儲存成功");
        }

        private void 儲存SToolStripButton1_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.uSERSBindingSource.EndEdit();
            this.uSERSTableAdapter.Update(this.uSERS._USERS);

            this.uSERMENUSBindingSource.EndEdit();
            this.uSERMENUSTableAdapter.Update(this.uSERS.USERMENUS);

            this.uSERSSHIPBindingSource.EndEdit();
            this.uSERSSHIPTableAdapter.Update(this.uSERS.USERSSHIP);

            MessageBox.Show("存檔成功");
        }

        private void button10_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = GetSCE();
        }

       



        public static void Truncate(string aa)
        {
            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            SqlCommand command = new SqlCommand("delete CUVV where IndexID=@aa ", MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@aa", aa));

            try
            {
                MyConnection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                MyConnection.Close();
            }
        }


        public static void Add(int IndexID, string Value, int LineNum)
        {
            SqlConnection MyConnection;

            MyConnection = globals.shipConnection;
            SqlCommand command = new SqlCommand("Insert into CUVV(IndexID,Value,LineNum) values(@IndexID,@Value,@LineNum)", MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@IndexID", IndexID));
            command.Parameters.Add(new SqlParameter("@Value", Value));
            command.Parameters.Add(new SqlParameter("@LineNum", LineNum));

            try
            {
                MyConnection.Open();
                command.ExecuteNonQuery();
            }
            finally
            {
                MyConnection.Close();
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = GetCUVV();
        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            try
            {
                Truncate(textBox7.Text);
                if (dataGridView3.Rows.Count > 0)
                {

                    for (int i = 0; i <= dataGridView3.Rows.Count - 1; i++)
                    {
                        DataGridViewRow row;

                        row = dataGridView3.Rows[i];

                        string T1 = row.Cells["VALUE"].Value.ToString();
                        string T2 = row.Cells["LINENUM"].Value.ToString();
                        if (T1 != "")
                        {
                            Add(Convert.ToInt16(textBox7.Text), T1, Convert.ToInt32(T2));
                        }
                    }
                }
            }
            catch { 
            }
        }

        private void toolStripButton8_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aCME_ARES_MAILBindingSource.EndEdit();
            this.aCME_ARES_MAILTableAdapter.Update(this.mail.ACME_ARES_MAIL);

            MessageBox.Show("儲存成功");
        }

        private void toolStripButton9_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aCME_AUTOPROBindingSource.EndEdit();
            this.aCME_AUTOPROTableAdapter.Update(this.uSERS.ACME_AUTOPRO);

            MessageBox.Show("儲存成功");
        }

        private void button14_Click(object sender, EventArgs e)
        {
            textBox3.Text = "";
        }

        private void button11_Click(object sender, EventArgs e)
        {
            int ff = Convert.ToInt32(textBox9.Text);
            AddAUOGD6GB();
            DateTime dt = new DateTime(ff, 1, 1);
            for (int i = 0; i <= 365; i++)
            {
                string f = "";


                if ((dt.AddDays(i).DayOfWeek.ToString() == "Monday"))
    
                {
                    DateTime f1 = dt.AddDays(i);
                    string DATE = f1.ToString("yyyyMMdd");
                    string DATE3 = f1.AddDays(-1).ToString("yyyyMMdd");
                    int DAY = f1.Day;
                    string STARTDATE = f1.ToString("yyyyMM") + "01";
                    string STARTDATE2 = f1.AddMonths(-1).ToString("yyyyMM") + "01";
                    DateTime f2 = Convert.ToDateTime(f1.ToString("yyyy") + "/" + f1.ToString("MM") + "/01");
                    string DATE2 = f2.AddDays(-1).ToString("yyyyMMdd");

                    int a = NumOfWeek(f1);
                    if (a > 1 && DAY > 7)
                    {
                        AddAUOGD5GB(DATE, STARTDATE, DATE3);
                    }
                    else
                    {
                        AddAUOGD5GB(DATE, STARTDATE2, DATE2); 
                    }
             
                }

            }
            MessageBox.Show("新增成功");
        }


        int NumOfWeek(DateTime aDt)
        {
            int nDay = aDt.Day;
            int nDayOfWeek = aDt.DayOfWeek - DayOfWeek.Sunday;
            int nRemainder = (nDayOfWeek - nDay + 1) % 7;
            if (nRemainder < 0) { nRemainder = 7 + nRemainder; }
            return (6 + nDay + nRemainder) / 7;

        }

        private void button15_Click(object sender, EventArgs e)
        {
            int ff = Convert.ToInt32(textBox9.Text);
            DateTime dt = new DateTime(ff, 1, 1);
            int T = 0;
            for (int i = 0; i <= 365; i++)
            {
               
                if ((dt.AddDays(i).DayOfWeek.ToString() == "Sunday"))
                {
                    T++;
                    DateTime f1 = dt.AddDays(i);
                    string DATE = f1.ToString("yyyyMMdd");
                    string DATE3 = f1.AddDays(+6).ToString("yyyyMMdd");

                    ADDGBWEEK(DATE, DATE3, T);

                }

            }
            MessageBox.Show("新增成功");
        }

     

        private void button17_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.pARAMSBindingSource.EndEdit();
            this.pARAMSTableAdapter.Update(this.uSERS.PARAMS);

            MessageBox.Show("儲存成功");
        }

        private void button18_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\毛利.xlsx";

            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //產生 Excel Report
            ExcelReport.ExcelReportE1(GETORDER(), ExcelTemplate, OutPutFile, "pivot");
        }

        private System.Data.DataTable GETORDER()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT '2015' 年,MONTH(DDATE) 月,");
            sb.Append(" 			  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,");
            sb.Append(" 			   SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62015 ");
            sb.Append("                WHERE CARDGROUP=103");
            if (checkBox1.Checked)
            {
                sb.Append("          AND CARDCODE NOT IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')     ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("          AND CARDCODE  IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            sb.Append("           GROUP BY MONTH(DDATE)  ");
            sb.Append("                UNION ALL ");
            sb.Append("               SELECT '2016',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62016 ");
            sb.Append("                WHERE CARDGROUP=103");
            if (checkBox1.Checked)
            {
                sb.Append("          AND CARDCODE NOT IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("          AND CARDCODE  IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            sb.Append("           GROUP BY MONTH(DDATE)  ");
            sb.Append("                UNION ALL ");
            sb.Append("               SELECT '2017',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62017 ");
            sb.Append("                WHERE CARDGROUP=103");
            if (checkBox1.Checked)
            {
                sb.Append("          AND CARDCODE NOT IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("          AND CARDCODE  IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            sb.Append("           GROUP BY MONTH(DDATE)  ");
            sb.Append("                 UNION ALL ");
            sb.Append("               SELECT '2018',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62018 ");
            sb.Append("                WHERE CARDGROUP=103");
            if (checkBox1.Checked)
            {
                sb.Append("          AND CARDCODE NOT IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("          AND CARDCODE IN  ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            sb.Append("           GROUP BY MONTH(DDATE)  ");
            sb.Append("                 UNION ALL ");
            sb.Append("               SELECT '2019',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62019 ");
            sb.Append("                WHERE CARDGROUP=103");
            if (checkBox1.Checked)
            {
                sb.Append("          AND CARDCODE NOT IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("          AND CARDCODE IN  ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            sb.Append("           GROUP BY MONTH(DDATE)  ");
            sb.Append("                 UNION ALL ");
            sb.Append("               SELECT '2020',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62020 ");
            sb.Append("                WHERE CARDGROUP=103");
            if (checkBox1.Checked)
            {
                sb.Append("          AND CARDCODE NOT IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("          AND CARDCODE IN  ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            sb.Append("           GROUP BY MONTH(DDATE)  ");
            sb.Append("                 UNION ALL ");
            sb.Append("               SELECT '2021',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62021 ");
            sb.Append("                WHERE CARDGROUP=103");
            if (checkBox1.Checked)
            {
                sb.Append("          AND CARDCODE NOT IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            if (checkBox2.Checked)
            {
                sb.Append("          AND CARDCODE  IN ('''1349-00','''0257-00','''0511-00','''0060-00','''0021-00','''1030-00')    ");
            }
            sb.Append("           GROUP BY MONTH(DDATE)  ");
            sb.Append("                 ORDER BY 年,CAST(MONTH(DDATE) AS INT) ");
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
        private System.Data.DataTable GETORDER2S()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("	SELECT 年, Q,CAST(SUM(收入) AS decimal) 收入,CAST(SUM(成本)*-1 AS decimal) 成本,CAST(SUM(費用)*-1 AS decimal) 費用");
            sb.Append("	,CAST(SUM(收入)+SUM(成本)+SUM(費用) AS decimal) 淨利,ROUND((SUM(收入)+SUM(成本)+SUM(費用))/SUM(收入),4)*100 百分比  FROM (	");
            sb.Append("			 	 	SELECT YEAR(REFDATE) 年,");
            sb.Append("				(CASE WHEN MONTH(REFDATE) BETWEEN 1 AND 3 THEN 'Q1'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 4 AND 6 THEN 'Q2'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 10 AND 12 THEN 'Q4' END) Q,  SUM(CREDIT)-SUM(DEBIT) 收入,0 成本,0 費用");
            sb.Append("			 FROM JDT1 WHERE   SUBSTRING(ACCOUNT,1,2) between '41' and '42' ");
            sb.Append("			 GROUP BY ");
            sb.Append("			 YEAR(REFDATE) ,");
            sb.Append("				(CASE WHEN MONTH(REFDATE) BETWEEN 1 AND 3 THEN 'Q1'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 4 AND 6 THEN 'Q2'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 10 AND 12 THEN 'Q4' END) ");
            sb.Append("			 UNION ALL");
            sb.Append("			SELECT (YEAR(REFDATE)) 年,	(CASE WHEN MONTH(REFDATE) BETWEEN 1 AND 3 THEN 'Q1'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 4 AND 6 THEN 'Q2'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 10 AND 12 THEN 'Q4' END) Q,0,SUM(CREDIT)-SUM(DEBIT) AMOUNT,0");
            sb.Append("			 FROM JDT1 ");
            sb.Append("			 WHERE  SUBSTRING(ACCOUNT,1,1) = '5'  ");
            sb.Append("			 			 GROUP BY ");
            sb.Append("			 YEAR(REFDATE) ,");
            sb.Append("				(CASE WHEN MONTH(REFDATE) BETWEEN 1 AND 3 THEN 'Q1'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 4 AND 6 THEN 'Q2'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 10 AND 12 THEN 'Q4' END) ");
            sb.Append("							 UNION ALL");
            sb.Append("			SELECT (YEAR(REFDATE)) 年,	(CASE WHEN MONTH(REFDATE) BETWEEN 1 AND 3 THEN 'Q1'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 4 AND 6 THEN 'Q2'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 10 AND 12 THEN 'Q4' END) Q,0,0,SUM(CREDIT)-SUM(DEBIT) AMOUNT");
            sb.Append("			 FROM JDT1 ");
            sb.Append("			 WHERE  SUBSTRING(ACCOUNT,1,1) ='6'");
            sb.Append("			 			 GROUP BY ");
            sb.Append("			 YEAR(REFDATE) ,");
            sb.Append("				(CASE WHEN MONTH(REFDATE) BETWEEN 1 AND 3 THEN 'Q1'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 4 AND 6 THEN 'Q2'");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append("				WHEN MONTH(REFDATE) BETWEEN 10 AND 12 THEN 'Q4' END) ");
            sb.Append(") AS A GROUP BY 年,Q ORDER BY  年, Q");
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

        private System.Data.DataTable GETORDER2()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT '2015' 年,MONTH(DDATE) 月,");
            sb.Append(" 			  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,");
            sb.Append(" 			   SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62015 ");
            sb.Append("                WHERE CARDGROUP=116    GROUP BY MONTH(DDATE)  ");
            sb.Append("                UNION ALL ");
            sb.Append("               SELECT '2016',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62016 ");
            sb.Append("                WHERE CARDGROUP=116   GROUP BY MONTH(DDATE) ");
            sb.Append("                UNION ALL ");
            sb.Append("               SELECT '2017',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62017 ");
            sb.Append("                WHERE CARDGROUP=116  GROUP BY MONTH(DDATE)  ");
            sb.Append("                 UNION ALL ");
            sb.Append("               SELECT '2018',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62018 ");
            sb.Append("                WHERE CARDGROUP=116  GROUP BY MONTH(DDATE)  ");
                        sb.Append("                 UNION ALL ");
            sb.Append("               SELECT '2019',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62019 ");
            sb.Append("                WHERE CARDGROUP=116  GROUP BY MONTH(DDATE)  ");
            sb.Append("                 UNION ALL ");
            sb.Append("               SELECT '2020',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62020 ");
            sb.Append("                WHERE CARDGROUP=116  GROUP BY MONTH(DDATE)  ");
            sb.Append("                 UNION ALL ");
            sb.Append("               SELECT '2021',MONTH(DDATE) 月,");
            sb.Append(" 			  		  CASE WHEN　MONTH(DDATE) BETWEEN 1 AND 3 THEN 'Q1' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 4 AND 6 THEN 'Q2' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 7 AND 9 THEN 'Q3' ");
            sb.Append(" 			  WHEN　MONTH(DDATE) BETWEEN 10 AND 12 THEN 'Q4' END Q,SUM(GTOTAL) 金額,SUM(GTOTAL)-SUM(GVALUE) 毛利,ROUND((SUM(GTOTAL)-SUM(GVALUE))/SUM(GTOTAL),4)*100 毛利率 FROM Account_Temp62021 ");
            sb.Append("                WHERE CARDGROUP=116  GROUP BY MONTH(DDATE)  ");
            sb.Append("                 ORDER BY 年,CAST(MONTH(DDATE) AS INT) ");
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

        private void button19_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\毛利.xlsx";

            //Excel的樣版檔
            string ExcelTemplate = FileName;

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            //產生 Excel Report
            ExcelReport.ExcelReportE1(GETORDER2(), ExcelTemplate, OutPutFile, "pivot");
        }

        private void fillToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                this.y_2004TableAdapter.Fill(this.uSERS.Y_2004, ((decimal)(System.Convert.ChangeType(yEARToolStripTextBox.Text, typeof(decimal)))));
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void toolStripButton10_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.y_2004BindingSource.EndEdit();
            this.y_2004TableAdapter.Update(this.uSERS.Y_2004);

            MessageBox.Show("儲存成功");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.rMA_PARAMSBindingSource.EndEdit();
            this.rMA_PARAMSTableAdapter.Update(this.uSERS.RMA_PARAMS);

            MessageBox.Show("儲存成功");
        }



        private void button21_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                string F = opdf.FileName;
                GetExcelContentGD4(F);

            }
        }
        private void GetExcelContentGD4(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string COUNTRY = "";
            string LOCATION = "";

            // int u = 0;
            int v = 0;
            for (int b = 1; b <= 100; b++)
            {


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, b]);
                range.Select();
                id = range.Text.ToString().Trim();

                if (id == "國家")
                {

                    //u = 5;
                    v = b;
                    // break;

                    for (int i = 1; i <= 100; i++)
                    {


                        //for (int j = u; j <= L1; j++)
                        //{
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1 + i, b]);
                        range.Select();
                        COUNTRY = range.Text.ToString().Trim();


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1 + i, b + 1]);
                        range.Select();
                        LOCATION = range.Text.ToString().Trim();


                        if (!String.IsNullOrEmpty(COUNTRY))
                        {
                            if (!String.IsNullOrEmpty(LOCATION))
                            {
                                AddAUOGD4("出口", COUNTRY, LOCATION);
                            }
                        }

                    }


                }
            }


            //if (u == 0)
            //{
            //    MessageBox.Show("Excel格式有誤");
            //    return;

            //}




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
            MessageBox.Show("匯出成功");
        }

        private void GetExcelF(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string KG;
            string LOCATION = "";
            string FEE = "";




            for (int i = 2; i <= 100; i++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                KG = range.Text.ToString().Trim();
                if (!KG.Contains(".")) 
                {
                    KG += ".0";
                }
                for (int b = 2; b <= 9; b++)
                {
                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, b]);
                    range.Select();
                    LOCATION = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, b]);
                    range.Select();
                    FEE = range.Text.ToString().Replace(",", "").Trim();

                    if (!String.IsNullOrEmpty(FEE) && !String.IsNullOrEmpty(KG) && !String.IsNullOrEmpty(LOCATION))
                    {

                        AddF(comboBox9.Text, KG, LOCATION, FEE);

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
            System.GC.WaitForPendingFinalizers();
            MessageBox.Show("匯出成功");
        }
        private void GetExcelP(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string PFROM;
            string LOCATION = "";
            string PTO = "";



            //直
            for (int i = 2; i <= 11; i++)
            {

                //橫
                for (int b = 2; b <= 11; b++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                    range.Select();
                    PFROM = range.Text.ToString().Replace(",", "").Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[1, b]);
                    range.Select();
                    PTO = range.Text.ToString().Replace(",", "").Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, b]);
                    range.Select();
                    LOCATION = range.Text.ToString().Replace(",", "").Trim();

                    if (!String.IsNullOrEmpty(PFROM) && !String.IsNullOrEmpty(PTO) && !String.IsNullOrEmpty(LOCATION))
                    {

                        AddP(PFROM, PTO, LOCATION);

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
            System.GC.WaitForPendingFinalizers();
            MessageBox.Show("匯出成功");
        }
        public void AddAUOGD4(string CTYPE, string COUNTRY, string LOCATION)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_DHL_COUNTRY(CTYPE,COUNTRY,LOCATION) values(@CTYPE,@COUNTRY,@LOCATION)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CTYPE", CTYPE));
            command.Parameters.Add(new SqlParameter("@COUNTRY", COUNTRY));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));

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
        public void DELF(string CTYPE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("DELETE WH_DHL_FEE WHERE CTYPE=@CTYPE ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CTYPE", CTYPE));

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

        public void AddF(string CTYPE, string KG, string LOCATION, string FEE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_DHL_FEE(CTYPE,KG,LOCATION,FEE) values(@CTYPE,@KG,@LOCATION,@FEE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CTYPE", CTYPE));
            command.Parameters.Add(new SqlParameter("@KG", KG));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
            command.Parameters.Add(new SqlParameter("@FEE", FEE));
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

        public void AddP(string PFROM, string PTO, string LOCATION)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into WH_DHL_PART(PFROM,PTO,LOCATION) values(@PFROM,@PTO,@LOCATION)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@PFROM", PFROM));
            command.Parameters.Add(new SqlParameter("@PTO", PTO));
            command.Parameters.Add(new SqlParameter("@LOCATION", LOCATION));
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

        public void UPDATEE(string ENGNAME, string COUNTRY)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("UPDATE WH_DHL_COUNTRY SET ENGNAME=@ENGNAME WHERE COUNTRY =@COUNTRY", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ENGNAME", ENGNAME));
            command.Parameters.Add(new SqlParameter("@COUNTRY", COUNTRY));
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
        private void GetExcelContentGD44(string ExcelFile, int Y)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            //     int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            //    Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string id;
            string id2 = "";
            string id3 = "";
            string id4 = "";
            string idG = "";

            int u = 0;
            int v = 0;
            int L1 = 0;


            for (int b = 5; b <= 20; b++)
            {
                for (int jj = 1; jj <= 20; jj++)
                {

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[jj, b]);
                    range.Select();
                    id = range.Text.ToString().Trim();




                    if (id == "金额")
                    {
                        u = jj + 1;
                        v = b + 1;
                        break;
                    }

                }

            }
            for (int U = u + 2; U <= 1000; U++)
            {
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[U, 1]);
                range.Select();
                id = range.Text.ToString().Trim();
                if (String.IsNullOrEmpty(id))
                {
                    L1 = U - 1;
                    break;
                }

            }

            if (u == 0)
            {
                MessageBox.Show("Excel格式有誤");
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
                return;

            }
            for (int i = v; i <= Y; i++)
            {




                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[u, i]);
                range.Select();
                id = range.Text.ToString().Trim();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[u - 1, i]);
                range.Select();
                idG = range.Text.ToString().Trim();



                //try
                //{


                if (id != "车型")
                {



                    for (int j = u; j <= L1; j++)
                    {
                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, i]);
                        range.Select();
                        id3 = range.Text.ToString().Trim();


                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, 5]);
                        range.Select();

                        id4 = range.Text.ToString().Trim();



                        range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[j, 3]);
                        range.Select();
                        id2 = range.Text.ToString().Trim();

                        int FF = id4.IndexOf("SH");
                        if (FF.ToString() != "-1")
                        {

                            if ((!String.IsNullOrEmpty(id4)) && id3.Trim() != "" && id3.Trim() != "0" && id3.Trim() != "0.00" && id3.Trim() != "/")
                            {
                                string hj = "";
                                if (comboBox2.Text != "蘇州宏高")
                                {
                                    hj = "";
                                }
                                else
                                {
                                    hj = comboBox3.Text;
                                }
                                decimal n;
                                //if (decimal.TryParse(id3, out n))
                                //{
                                //    decimal cd = Convert.ToDecimal(id3) * Convert.ToDecimal(textBox1.Text);
                                //    //if (cd == -1490)
                                //    //{
                                //    //    MessageBox.Show("A");
                                //    //}
                                //    if (idG.Trim() != "合计" && idG.Trim().ToUpper() != "TOTAL" && !String.IsNullOrEmpty(id))
                                //    {
                                //        if (fmLogin.LoginID.ToString().ToUpper() != "LLEYTONCHEN")
                                //        {
                                //            AddAUOGD4(id4, id, cd.ToString(), comboBox2.Text, comboBox3.Text, id2, comboBox1.Text, textBox1.Text, id3, comboBox6.Text);
                                //        }
                                //    }
                                //    else
                                //    {
                                //        if (cd < 0)
                                //        {
                                //            if (fmLogin.LoginID.ToString().ToUpper() != "LLEYTONCHEN")
                                //            {
                                //                AddAUOGD4(id4, "", cd.ToString(), comboBox2.Text, comboBox3.Text, id2, comboBox1.Text, textBox1.Text, id3, comboBox6.Text);
                                //            }
                                //        }
                                //    }
                                //}

                            }
                        }


                    }


                }



                //}

                //catch (Exception ex)
                //{
                //    MessageBox.Show(ex.Message);
                //}
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
            MessageBox.Show("匯出成功");
        }

        private void button22_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                DELF(comboBox9.Text);
                string F = opdf.FileName;
                GetExcelF(F);

            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                string F = opdf.FileName;
                GetExcelP(F);

            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                string F = opdf.FileName;
                GetExcelE(F);

            }
        }
        private void GetExcelE(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = true;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            string excelFile = ExcelFile;

            //Open the worksheet file
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);

            int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;

            int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;

            Hashtable ht = new Hashtable(iRowCnt);



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);


            string ENAME;
            string NAME = "";
      
            for (int i = 1; i <= 250; i++)
            {

                

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                    range.Select();
                    ENAME = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                    range.Select();
                    NAME = range.Text.ToString().Trim();

                    if (!String.IsNullOrEmpty(ENAME) && !String.IsNullOrEmpty(NAME))
                    {

                        UPDATEE(ENAME, NAME);

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
            MessageBox.Show("匯出成功");
        }

        private void fillToolStripButton1_Click(object sender, EventArgs e)
        {
            try
            {
                this.sATT21TableAdapter.Fill(this.sa.SATT21, mEMOToolStripTextBox.Text);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

        }

        private void button25_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sATT21BindingSource.EndEdit();
            this.sATT21TableAdapter.Update(this.sa.SATT21);

            MessageBox.Show("儲存成功");
        }

        private void button26_Click(object sender, EventArgs e)
        {

        }

        private void button26_Click_1(object sender, EventArgs e)
        {

        }

 

        private void button27_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aCME_MAIL_BACKUP2BindingSource.EndEdit();
            this.aCME_MAIL_BACKUP2TableAdapter.Update(this.mail.ACME_MAIL_BACKUP2);

            MessageBox.Show("儲存成功");
        }

        private void button28_Click(object sender, EventArgs e)
        {
            try
            {
                this.aCME_MAIL_BACKUP2TableAdapter.Fill(this.mail.ACME_MAIL_BACKUP2, ((int)(System.Convert.ChangeType(textBox10.Text, typeof(int)))));
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private void button26_Click_2(object sender, EventArgs e)
        {

        }

        private void button29_Click(object sender, EventArgs e)
        {

            // TODO: 這行程式碼會將資料載入 'uSERS.RPA_PackingH' 資料表。您可以視需要進行移動或移除。
            this.rPA_PackingHTableAdapter.Fill(this.uSERS.RPA_PackingH, textBox13.Text);

            this.rPA_PackingDTableAdapter.Fill(this.uSERS.RPA_PackingD);

        }

        private void button30_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.rPA_PackingHBindingSource.EndEdit();
            this.rPA_PackingHTableAdapter.Update(this.uSERS.RPA_PackingH);

            this.rPA_PackingDBindingSource.EndEdit();
            this.rPA_PackingDTableAdapter.Update(this.uSERS.RPA_PackingD);
        }

        private void button31_Click(object sender, EventArgs e)
        {
            UPWH_FEE();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            DA();
           
        }


        private void DA()
        { 
           System.Data.DataTable dtCost = MakeTableCombine();

           StringBuilder sb = new StringBuilder();
            DataRow dr = null;
            int DD = 0;
            System.Data.DataTable dt1 = GetDataA();
            for (int j = 0; j <= dt1.Rows.Count - 1; j++)
            {
                string ID = dt1.Rows[j]["ID"].ToString();
                string DOCENTRY = dt1.Rows[j]["DOCENTRY"].ToString().Trim();
                System.Data.DataTable dt = GetDataE(ID,DOCENTRY);
                sb.Append("" + DOCENTRY + ",");
                for (int i = 0; i <= dt.Rows.Count - 1; i++)
                {
                    DD++;
                    DataRow dd = dt.Rows[i];
                    dr = dtCost.NewRow();
                    dr["DD"] = DD;
                    dr["廠商編號"] = dd["廠商編號"].ToString();
                    dr["廠商名稱"] = dd["廠商名稱"].ToString();
                    dr["統一編號"] = dd["統一編號"].ToString();
                    dr["付款條件"] = dd["付款條件"].ToString();
                    dr["滙費負擔"] = dd["滙費負擔"].ToString();
                    dr["MBU"] = dd["MBU"].ToString();
                    dr["申請日期"] = dd["申請日期"].ToString();
                    dr["付款日期"] = dd["付款日期"].ToString();
                    dr["採購日期"] = dd["採購日期"].ToString();
                    dr["採購單號碼"] = dd["採購單號碼"].ToString();
                    dr["工單號碼"] = dd["工單號碼"].ToString();


                    dr["歸屬部門"] = dd["歸屬部門"].ToString();
                    dr["科目代碼"] = dd["科目代碼"].ToString();
                    dr["費用名稱"] = dd["費用名稱"].ToString();
                    dr["未稅金額"] = dd["未稅金額"].ToString();
                    dr["稅額"] = dd["稅額"].ToString();
                    dr["加總"] = dd["加總"].ToString();

                    dr["T含稅金額"] = dd["T含稅金額"].ToString();
                    dr["T未稅金額"] = dd["T未稅金額"].ToString();
                    dr["T稅額"] = dd["T稅額"].ToString();
                    dr["使用者"] = dd["使用者"].ToString();
                    dr["ID"] = dd["ID"].ToString();
                    dtCost.Rows.Add(dr);

                }
                System.Data.DataTable dt2 = GetDataE2(ID, DOCENTRY);
                if (dt2.Rows.Count > 0)
                {
                    dr = dtCost.NewRow();
                    DD++;
                    dr["DD"] = DD;
                    dr["廠商編號"] = "";
                    dr["廠商名稱"] = "";
                    dr["統一編號"] = "";
                    dr["付款條件"] = "";
                    dr["滙費負擔"] = "";

                    dr["申請日期"] = "";
                    dr["付款日期"] = "";
                    dr["採購日期"] = "";
                    dr["採購單號碼"] = "";
                    dr["工單號碼"] = "";


                    dr["歸屬部門"] = "";
                    dr["科目代碼"] = "";
                    dr["費用名稱"] = "小計";
                    string ff = dt2.Rows[0]["未稅金額"].ToString(); 
                    dr["未稅金額"] = dt2.Rows[0]["未稅金額"].ToString();
                    dr["稅額"] = dt2.Rows[0]["稅額"].ToString();
                    dr["加總"] = dt2.Rows[0]["加總"].ToString();

                    dr["T含稅金額"] = "";
                    dr["T未稅金額"] = "";
                    dr["T稅額"] = "";
                    dr["使用者"] = "";
                    dr["ID"] = "";
                    dtCost.Rows.Add(dr);
                }
            }

    //        dataGridView4.DataSource = dtCost;
            string B1 = "//acmew08r2ap//table//SIGN//MANAGER//";

            string B3 = "//acmew08r2ap//table//SIGN//USER//";


            string S2 = "//acmew08r2ap//table//EXCEL//";
            string FileName = string.Empty;

            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = S2 + "ACME支付通知單2.xls";

            //  t1 = GetDataE();


            string OutPutFile = lsAppDir + "\\" + "132-核.xls";
            sb.Remove(sb.Length - 1, 1);
            ExcelReport.ExcelReportOutputP(dtCost, FileName, OutPutFile, B1 + "JOJOHSU.JPG", B3 + "MICHELLEKO" + ".JPG", 135, 160, 380, 160, PackOPCH2(sb.ToString()));
        }
            //        sb.Append(" SELECT * FROM (SELECT T0.MBU,T0.CARDCODE 廠商編號,T0.CARDNAME 廠商名稱,T0.SERNO 統一編號,T0.PAY3 付款條件,T0.PAY2 滙費負擔,T0.DOCDATE 申請日期,T0.PREDATE 付款日期    ");
            //sb.Append(" ,T1.DOCDATE 採購日期,T1.DOCENTRY 	採購單號碼,T1.SHIPNO 工單號碼,T1.BU 歸屬部門,T1.ACCOUNT 科目代碼 ");
            //sb.Append(" ,T1.Dscription 費用名稱,T1.TOTAL 未稅金額,T1.RATE 稅額,T1.AMOUNT 加總,T0.AMOUNT T含稅金額,T0.ATOTAL T未稅金額,T0.ARATE T稅額,UserSign 使用者,T0.ID  FROM ACMESQLEEP.DBO.ACME_OITT T0 ");
        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("DD", typeof(int));
            dt.Columns.Add("MBU", typeof(string));
            dt.Columns.Add("廠商編號", typeof(string));
            dt.Columns.Add("廠商名稱", typeof(string));
            dt.Columns.Add("統一編號", typeof(string));
            dt.Columns.Add("付款條件", typeof(string));
            dt.Columns.Add("滙費負擔", typeof(string));
            dt.Columns.Add("申請日期", typeof(string));

            dt.Columns.Add("付款日期", typeof(string));
            dt.Columns.Add("採購日期", typeof(string));
            dt.Columns.Add("採購單號碼", typeof(string));
            dt.Columns.Add("工單號碼", typeof(string));
            dt.Columns.Add("歸屬部門", typeof(string));
            dt.Columns.Add("科目代碼", typeof(string));
            dt.Columns.Add("費用名稱", typeof(string));
            dt.Columns.Add("未稅金額", typeof(string));
            dt.Columns.Add("稅額", typeof(string));
            dt.Columns.Add("加總", typeof(string));
            dt.Columns.Add("T含稅金額", typeof(string));
            dt.Columns.Add("T未稅金額", typeof(string));
            dt.Columns.Add("T稅額", typeof(string));
            dt.Columns.Add("使用者", typeof(string));
            dt.Columns.Add("ID", typeof(string));

            return dt;
        }

        public static System.Data.DataTable PackOPCH2(string AA)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,ROUND(SUM(t1.totalsumsy),0) 未稅金額,ROUND(SUM(t1.linevat),0) 稅額,ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PCH1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ");
            sb.Append(" union all");
            sb.Append("  SELECT  '加總','' 科目名稱,ROUND(SUM(t1.totalsumsy),0) 未稅金額,ROUND(SUM(t1.linevat),0) 稅額,ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PCH1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }
                private System.Data.DataTable GetDataA()
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT DISTINCT ID,DOCENTRY FROM ACMESQLEEP.DBO.ACME_ITT1     WHERE ID='160'       ");

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
        private System.Data.DataTable GetSS(string  ITEMCODE)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT * FROM  ACMESQL05.DBO.OITM  WHERE ITEMCODE=@ITEMCODE      ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
               command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));



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
        private System.Data.DataTable GetDataE(string ID, string DOCENTRY)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT T0.MBU,T0.CARDCODE 廠商編號,T0.CARDNAME 廠商名稱,T0.SERNO 統一編號,T0.PAY3 付款條件,T0.PAY2 滙費負擔,T0.DOCDATE 申請日期,T0.PREDATE 付款日期     ");
            sb.Append(" ,T1.DOCDATE 採購日期,T1.DOCENTRY 	採購單號碼,T1.SHIPNO 工單號碼,T1.BU 歸屬部門,T1.ACCOUNT 科目代碼  ");
            sb.Append(" ,T1.Dscription 費用名稱,T1.TOTAL 未稅金額,T1.RATE 稅額,T1.AMOUNT 加總,T0.AMOUNT T含稅金額,T0.ATOTAL T未稅金額,T0.ARATE T稅額,UserSign 使用者,T0.ID  FROM ACMESQLEEP.DBO.ACME_OITT T0  ");
            sb.Append(" LEFT JOIN ACMESQLEEP.DBO.ACME_ITT1 T1 ON (T0.ID=T1.ID) ");
            sb.Append(" WHERE  T0.ID=@ID  AND T1.DOCENTRY=@DOCENTRY  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
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

        private System.Data.DataTable GetDataE2(string ID, string DOCENTRY)
        {

            SqlConnection connection = globals.Connection;

            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT SUM(T1.TOTAL) 未稅金額,SUM(T1.RATE) 稅額,SUM(T1.AMOUNT) 加總 FROM ACMESQLEEP.DBO.ACME_ITT1 T1");
            sb.Append(" WHERE  T1.ID=@ID  AND T1.DOCENTRY=@DOCENTRY ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));
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

        public static System.Data.DataTable GetDataE3(string AA)
        {

            SqlConnection MyConnection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();


            sb.Append("  SELECT  T1.ACCTCODE 科目,MAX(T2.ACCTNAME) 科目名稱,ROUND(SUM(t1.totalsumsy),0) totalsumsy,ROUND(SUM(t1.linevat),0) linevat,ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PCH1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ");
            sb.Append(" union all");
            sb.Append("   SELECT  '四捨五入差異','',0,0,SUM(t1.DOCTOTAL)-( SELECT SUM(加總) A FROM ( SELECT   ROUND(SUM(t1.totalsumsy+t1.linevat),0)  加總 FROM PCH1 T1 LEFT JOIN OACT  T2 ON (T1.ACCTCODE=T2.ACCTCODE) WHERE   t1.docentry  IN (" + AA + ") GROUP BY T1.ACCTCODE ) AS A");
            sb.Append(" ) FROM OPCH T1  WHERE   t1.docentry  IN (" + AA + ")  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);

            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, "OPDN");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["OPDN"];

        }

        public void DELOPOR()
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand(" TRUNCATE TABLE AP_TEMP1  ", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));

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

        private void button33_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
                DELOPOR();
                GD6(opdf.FileName);


                System.Data.DataTable G1 = GetOPOR();
                dataGridView1.DataSource = G1;



            }
        }

        private void GD6(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

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



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);



            for (int i = 2; i <= iRowCnt; i++)
            {


                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }
                string EMPID;
                string USER;
                string USERNAME;
                string GRADE;
                string VER;
                decimal QTY;
                decimal PRICE;
                decimal AMT;
                string REMARK1;
                string REMARK2;
                string P1;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                EMPID = range.Text.ToString().Trim();

                //if (ITEMCODE != "")
                //{
                //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                //    range.Select();
                //    LINE = range.Text.ToString().Trim().ToUpper();

                //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                //    range.Select();
                //    GRADE = range.Text.ToString().Trim().ToUpper();

                //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 3]);
                //    range.Select();
                //    VER = range.Text.ToString().Trim().ToUpper();

                //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 4]);
                //    range.Select();
                //    QTY = Convert.ToDecimal(range.Text.ToString().Trim());

                //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 5]);
                //    range.Select();
                //    P1 = range.Text.ToString().Trim();

                //    decimal n;
                //    if (decimal.TryParse(P1, out n))
                //    {
                //        PRICE = Convert.ToDecimal(P1);
                //    }
                //    else
                //    {
                //        PRICE = 0;
                //    }

                //    AMT = PRICE * QTY;

                //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 7]);
                //    range.Select();
                //    REMARK1 = range.Text.ToString().Trim();

                //    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 8]);
                //    range.Select();
                //    REMARK2 = range.Text.ToString().Trim();

                //    try
                //    {
                //        //if (!String.IsNullOrEmpty(ITEMCODE))
                //        //{

                //        //    string ITEM = "";
                      


                //        //    ADDOPOR(Convert.ToInt16(LINE), ITEM, QTY, PRICE, AMT, REMARK1 + " " + REMARK2, "S0001-GD", fmLogin.LoginID.ToString(), ITEMCODE + "." + VER.Substring(0, 1), "", "");
                //        //}


                //    }

                //    catch (Exception ex)
                //    {
                //        MessageBox.Show(ex.Message);
                //    }
                //}
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

        private void GD7(string ExcelFile)
        {

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);


            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

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



            Microsoft.Office.Interop.Excel.Range range = null;



            object SelectCell = "A1";
            range = excelSheet.get_Range(SelectCell, SelectCell);



            for (int i = 2; i <= iRowCnt; i++)
            {


                if (iRowCnt > 500)
                {
                    iRowCnt = 500;
                }
                string EMPID;
                string EMPID2;

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 1]);
                range.Select();
                EMPID = range.Text.ToString().Trim();
                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[i, 2]);
                range.Select();
                EMPID2 = range.Text.ToString().Trim();

                string OutPutFile = lsAppDir + "\\Excel\\temp2\\" +
   EMPID + EMPID2;

                if (!String.IsNullOrEmpty(EMPID))
                {
                    System.IO.Directory.CreateDirectory(OutPutFile);
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
        public void ADDOPOR(string S1, string S2, string S3)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into AP_TEMP1(S1,S2,S3) values(@S1,@S2,@S3)", connection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@S1", S1));
            command.Parameters.Add(new SqlParameter("@S2", S2));
            command.Parameters.Add(new SqlParameter("@S3", S3));

            //
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
        public static System.Data.DataTable GetOPOR()
        {
            SqlConnection MyConnection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  * FROM AP_TEMP1");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERS", fmLogin.LoginID.ToString()));
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

        private void button34_Click(object sender, EventArgs e)
        {
    

            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {
     
                GD7(opdf.FileName);


            }


        }

        private void button35_Click(object sender, EventArgs e)
        {
                           OpenFileDialog opdf = new OpenFileDialog();
                           DialogResult result = opdf.ShowDialog();
                           if (opdf.FileName.ToString() == "")
                           {
                               MessageBox.Show("請選擇檔案");
                           }
                           else
                           {
         
                               WriteExcelProduct6(opdf.FileName);


                           }
        }
        private void WriteExcelProduct6(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Range range = null;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

            oCompany = new SAPbobsCOM.Company();

            oCompany.Server = "acmesap";
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCompany.UseTrusted = false;
            oCompany.DbUserName = "sapdbo";
            oCompany.DbPassword = "@rmas";
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

            int i = 0; //  to be used as an index

            oCompany.CompanyDB = FA;
            oCompany.UserName = "manager";
            oCompany.Password = "0918";
            int result = oCompany.Connect();
            if (result == 0)
            {


                SAPbobsCOM.BusinessPartners oPURCH = null;
                oPURCH = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                //     
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
                excelSheet.Activate();

                int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



                string ENAME;

                string EENAME;
                string ECARD;

                string BU;
                string BU2;
                for (int iRecord = 2; iRecord <= iRowCnt; iRecord++)
                {


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 2]);
                    range.Select();
                    ENAME = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    EENAME = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    BU2 = range.Text.ToString().Trim();


                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    ECARD = range.Text.ToString().Trim();



                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    BU = range.Text.ToString().Trim();

                    oPURCH.CardCode = ECARD;
                    oPURCH.CardName = ENAME;
                    oPURCH.CardForeignName = EENAME;
                    oPURCH.CardType = SAPbobsCOM.BoCardTypes.cSupplier;
                    //oPURCH.   = "21700108";

                    //oPURCH.UserFields.Fields.Item("U_BU").Value = BU;
                    //oPURCH.UserFields.Fields.Item("U_BU2").Value = BU2;

                    int res = oPURCH.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
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
            System.GC.WaitForPendingFinalizers();




        }
        private void WDRS(string ExcelFile)
        {

            //Create an Excel App
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            excelApp.Visible = false;
            excelApp.DisplayAlerts = false;
            //Interop params
            object oMissing = System.Reflection.Missing.Value;

            //The Excel doc paths

            string excelFile = ExcelFile;

            Microsoft.Office.Interop.Excel.Range range = null;

            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(excelFile, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

            Microsoft.Office.Interop.Excel.Worksheet excelSheet = null;

            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

            oCompany = new SAPbobsCOM.Company();

            oCompany.Server = "acmesap";
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCompany.UseTrusted = false;
            oCompany.DbUserName = "sapdbo";
            oCompany.DbPassword = "@rmas";
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

            int i = 0; //  to be used as an index

            oCompany.CompanyDB = FA;
            oCompany.UserName = "manager";
            oCompany.Password = "19571215";
            int result = oCompany.Connect();
            if (result == 0)
            {


                string CARDCODE;

                string RMB;
                string ITEMCODE;
                string QTY;
                string PRICE;
                string WHS;
                string WORKDAY;
                string LDAY;
                string PDAY;
                string STATUS;
                SAPbobsCOM.Documents oORDR = null;
                oORDR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

                //     
                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets.get_Item(1);
                excelSheet.Activate();

                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 1]);
                range.Select();
                CARDCODE = range.Text.ToString().Trim();


                range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[2, 2]);
                range.Select();
                RMB = range.Text.ToString().Trim();

                oORDR.CardCode = "0123-00";
                oORDR.DocCurrency = "RMB";

                oORDR.SalesPersonCode = 26;
                oORDR.DocumentsOwner = 2;
                oORDR.DocRate = 1;
                oORDR.DocDueDate = DateTime.Now;
                int iRowCnt = excelSheet.UsedRange.Cells.Rows.Count;
                int iColCnt = excelSheet.UsedRange.Cells.Columns.Count;



                for (int iRecord = 2; iRecord <= iRowCnt-2; iRecord++)
                {




                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 3]);
                    range.Select();
                    ITEMCODE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 4]);
                    range.Select();
                    QTY = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 5]);
                    range.Select();
                    PRICE = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 7]);
                    range.Select();
                    WHS = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 8]);
                    range.Select();
                    WORKDAY = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 9]);
                    range.Select();
                    LDAY = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 10]);
                    range.Select();
                    PDAY = range.Text.ToString().Trim();

                    range = ((Microsoft.Office.Interop.Excel.Range)excelSheet.UsedRange.Cells[iRecord, 11]);
                    range.Select();
                    STATUS = range.Text.ToString().Trim();


                    if (ITEMCODE != "")
                    {
                        oORDR.Lines.ItemCode = ITEMCODE;
                        System.Data.DataTable G1 = GetSS(ITEMCODE);
                        if (G1.Rows.Count == 0)
                        {
                            MessageBox.Show(ITEMCODE);
                        }
                     //   oORDR.Lines.ItemCode = ITEMCODE;
                        oORDR.Lines.Quantity = Convert.ToDouble(QTY);
                        oORDR.Lines.WarehouseCode = "SZ001";
                        oORDR.Lines.Price = 0;
                        oORDR.Lines.VatGroup = "AR0%";
                        oORDR.Lines.UserFields.Fields.Item("U_ACME_WORKDAY").Value = "內銷";
                        oORDR.Lines.UserFields.Fields.Item("U_ACME_SHIPDAY").Value = "2020.11.20";
                        oORDR.Lines.UserFields.Fields.Item("U_ACME_WORK").Value = "2020.11.18";
                        oORDR.Lines.UserFields.Fields.Item("U_SHIPSTATUS").Value = STATUS;
                        oORDR.Lines.Add();
                    }

                }


                int res = oORDR.Add();
                if (res != 0)
                {
                    MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
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
        private void WDRS2()
        {



            SAPbobsCOM.Company oCompany = new SAPbobsCOM.Company();

            oCompany = new SAPbobsCOM.Company();

            oCompany.Server = "acmesap";
            oCompany.language = SAPbobsCOM.BoSuppLangs.ln_English;
            oCompany.UseTrusted = false;
            oCompany.DbUserName = "sapdbo";
            oCompany.DbPassword = "@rmas";
            oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;

            int i = 0; //  to be used as an index

            oCompany.CompanyDB = FA;
            oCompany.UserName = "manager";
            oCompany.Password = "0918";
            int result = oCompany.Connect();
            if (result == 0)
            {


                string CARDCODE;

              
                SAPbobsCOM.Documents oORDR = null;
                oORDR = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);



                oORDR.CardCode = "0001-00";
                   //         oORDR.DocCurrency = "NTD";

                      //      oORDR.SalesPersonCode = 143;
                           // oORDR.DocumentsOwner = 30;
                    //        oORDR.DocRate = 4.312;
                            oORDR.Lines.ItemCode = "ACME00001.00001";
                            oORDR.Lines.Quantity = 1;
                  //          oORDR.Lines.WarehouseCode = "s";
                            oORDR.Lines.Price = 0;
                        //    oORDR.Lines.UserFields.Fields.Item("U_ACME_WORKDAY").Value = "";
                       //     oORDR.Lines.UserFields.Fields.Item("U_ACME_SHIPDAY").Value = "2020.11.20";
                          //  oORDR.Lines.UserFields.Fields.Item("U_ACME_WORK").Value = "2020.11.18";
                           // oORDR.Lines.UserFields.Fields.Item("U_SHIPSTATUS").Value = STATUS;
                            oORDR.Lines.Add();

                


                    int res = oORDR.Add();
                    if (res != 0)
                    {
                        MessageBox.Show("上傳錯誤 " + oCompany.GetLastErrorDescription());
                    }
                 
            }















        }

        private void button36_Click(object sender, EventArgs e)
        {
            System.Data.DataTable G1 = GetD1();
            if (G1.Rows.Count > 0)
            {

                for (int i = 0; i <= G1.Rows.Count-1; i++)
                {
                    int D1 = 0;
                    int D2 = 0;
                    int D3 = 0;
                    string ID = G1.Rows[i][0].ToString();

                    DateTime H2 = Convert.ToDateTime(G1.Rows[i][1]);

                    if (H2.DayOfWeek.ToString() == "Saturday")
                    {
                        D1 = 0;
                        D2 = 1;
                        D3 = 0;

                    }
                    else if (H2.DayOfWeek.ToString() == "Sunday")
                    {
                        D1 = 0;
                        D2 = 0;
                        D3 = 1;

                    }
                    else
                    {
                        D1 = 1;
                        D2 = 0;
                        D3 = 0;
                    }
                    UPG5(ID, D1, D2, D3);
                }
            
            
            
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            OpenFileDialog opdf = new OpenFileDialog();
            DialogResult result = opdf.ShowDialog();
            if (opdf.FileName.ToString() == "")
            {
                MessageBox.Show("請選擇檔案");
            }
            else
            {

                 WDRS(opdf.FileName);

            //   WDRS2();
            }
        }

        private void button38_Click(object sender, EventArgs e)
        {
            textBox15.Text = textBox15.Text.ToUpper() + ".JPG";
        }

        private void button39_Click(object sender, EventArgs e)
        {
            if (textBox16.Text != "")
            {
                UPCHI16();
            }
        }
    }
}
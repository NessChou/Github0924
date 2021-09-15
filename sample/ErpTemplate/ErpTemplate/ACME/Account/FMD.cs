using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
namespace ACME
{
    public partial class FMD : Form
    {
        string strCn = "Data Source=acmesap;Initial Catalog=acmesqlsp;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public FMD()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox2.Text == "")
            {
                MessageBox.Show("請輸入日記帳憑證號碼");
                return;
            
            }
            if (IsNumber(textBox1.Text) == false)
            {
                MessageBox.Show("日記帳憑證號碼請輸入數字");
                return;
            }
            if (IsNumber(textBox2.Text) == false)
            {
                MessageBox.Show("交易號碼請輸入數字");
                return;
            }
            System.Data.DataTable T1 = Get3(textBox1.Text, textBox2.Text);

            if (T1.Rows.Count == 0)
            {
                MessageBox.Show("分錄沒有資料");
                return;
            }


            dataGridView1.DataSource = T1;
            System.Data.DataTable T2 = GetFMBG();

            for (int i = 0; i <= T2.Rows.Count - 1; i++)
            {
   
                DataRow  row = T2.Rows[i];
                string LINENUM = row["LINENUM"].ToString();
                decimal  DEBIT = Convert.ToDecimal(row["DEBIT"].ToString());
                AddAUOGD2F(LINENUM, DEBIT);
            }

            try
            {
                this.account_FMD1TableAdapter.Fill(this.accBank.Account_FMD1, textBox1.Text, textBox2.Text);
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        private System.Data.DataTable Get3(string BATCHNUM, string TRANSID)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  ACCOUNT 總帳科目,ACCTNAME 名稱,CAST(DEBIT as int) 借項,CAST(CREDIT AS INT) 貸項,LINEMEMO 摘要 FROM BTF1 T0");
            sb.Append(" LEFT JOIN OACT T1 ON (T0.ACCOUNT=T1.ACCTCODE)");
            sb.Append(" WHERE T0.BATCHNUM=@BATCHNUM AND T0.TRANSID=@TRANSID AND ACCOUNT=12640101");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BATCHNUM", BATCHNUM));
            command.Parameters.Add(new SqlParameter("@TRANSID", TRANSID));
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

        private System.Data.DataTable GetCOPY(string BATCHNUM)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT T0.BATCHNUM 日記帳憑證號碼,CASE STATUS WHEN 'O' THEN '未結' WHEN 'C' THEN '已結' END 狀態  ");
            sb.Append("               ,CONVERT(VARCHAR(8),DATEID,112) 日期 ,CAST(T0.LOCTOTAL AS INT)  總計,T1.USERSIGN,T1.REF2 製單人,T1.FinncPriod  FROM OBTD T0");
            sb.Append("              LEFT JOIN OBTF T1  ON (T0.BATCHNUM=T1.BATCHNUM)  WHERE T0.BATCHNUM=@BATCHNUM  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BATCHNUM", BATCHNUM));

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
        private System.Data.DataTable GetCOPY2(string BATCHNUM)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT LINE_ID LINE,ACCOUNT 科目,DEBIT 借項,CREDIT 貸項,VATGROUP 稅群組,LINEMEMO 摘要,PROJECT 專案,PROFITCODE 部門,TRANSTYPE,VATLINE ,DEBCRED FROM BTF1 where BATCHNUM=@BATCHNUM  ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BATCHNUM", BATCHNUM));

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
  
        private System.Data.DataTable GETFMDMAX()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT CASE ISNULL(MAX(LINEID),'') WHEN '' THEN 1 ELSE MAX(LINEID)+1 END LINEID   FROM ACMESQLSP.DBO.Account_FMD1 WHERE BATCHNUM=@BATCHNUM AND TRANSID=@TRANSID");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BATCHNUM", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@TRANSID", textBox2.Text));
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
        private System.Data.DataTable Get3SUM()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(CAST(DEBIT as int))  FROM BTF1 T0");
            sb.Append(" WHERE T0.BATCHNUM=@BATCHNUM AND T0.TRANSID=@TRANSID AND ACCOUNT=12640101");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BATCHNUM", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@TRANSID", textBox2.Text));
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
        private System.Data.DataTable GetSUMOJDT()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(CAST(DEBIT as int))  FROM JDT1 T0");
            sb.Append(" WHERE  T0.TRANSID=@TRANSID AND ACCOUNT=12640101");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@TRANSID", textBox3.Text));
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
        private System.Data.DataTable GetONNM()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT AUTOKEY NUM,AUTOKEY NUM1 FROM ONNM WHERE OBJECTCODE='FMD'");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
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

        private System.Data.DataTable GetPEROID()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT  MAX(PERIOD) PERIOD FROM DBO.[@CADMEN_FMD]");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
          
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
        private System.Data.DataTable GeOJDT()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT U_BSREN from dbo.[@CADMEN_FMD] WHERE U_BSREN=@U_BSREN");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_BSREN", textBox3.Text));
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

        private System.Data.DataTable GetFMB()
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT * FROM Account_FMD1 WHERE BATCHNUM=@BATCHNUM AND TRANSID=@TRANSID");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BATCHNUM", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@TRANSID", textBox2.Text));
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

        private System.Data.DataTable GetFMBG()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT LINE_ID+1 LINENUM,DEBIT FROM BTF1 T0");
            sb.Append(" WHERE BATCHNUM=@BATCHNUM AND TRANSID=@TRANSID AND ACCOUNT='12640101'");
            sb.Append(" AND LINE_ID+1 NOT IN (SELECT ISNULL(LINENUM,'') FROM ACMESQLSP.DBO.Account_FMD1 WHERE BATCHNUM=@BATCHNUM AND TRANSID=@TRANSID)");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BATCHNUM", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@TRANSID", textBox2.Text));
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
        private System.Data.DataTable GetFMBSUM()
        {

            SqlConnection connection = new SqlConnection(strCn);

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT CAST(SUM(U_PC_BSTAX) AS INT) AA FROM Account_FMD1 WHERE BATCHNUM=@BATCHNUM AND TRANSID=@TRANSID");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BATCHNUM", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@TRANSID", textBox2.Text));
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
        private void FMD_Load(object sender, EventArgs e)
        {
         
        }

        private void account_FMD1BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.account_FMD1BindingSource.EndEdit();
            this.account_FMD1TableAdapter.Update(this.accBank.Account_FMD1);

        }



        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= account_FMD1DataGridView.Rows.Count - 2; i++)
            {

                DataGridViewRow row;

                row = account_FMD1DataGridView.Rows[i];
                string D1 = row.Cells["U_PC_BSAPP"].Value.ToString();
                string D3 = row.Cells["U_PC_BSDAT"].Value.ToString();
                string D4 = row.Cells["U_PC_BSTAX"].Value.ToString();
                string D5 = row.Cells["U_PC_BSAMN"].Value.ToString();
                string D6 = row.Cells["U_PC_BSAMT"].Value.ToString();
                if (D1.Length != 8 || D3.Length != 8)
                {
                    MessageBox.Show("日期格式不符");
                    return;
                }
                if (String.IsNullOrEmpty(D4))
                {
                    MessageBox.Show("稅額不可空白");
                    return;
                }

                try
                {
                    decimal DD4 = Convert.ToDecimal(D4);
                    decimal DD5 = Convert.ToDecimal(D5);
                    decimal DD6 = Convert.ToDecimal(D6);
                }
                catch
                {
                    MessageBox.Show("金額必須為數字");
                    return;
                }
            }

            this.Validate();
            this.account_FMD1BindingSource.EndEdit();
            this.account_FMD1TableAdapter.Update(this.accBank.Account_FMD1);
            MessageBox.Show("存檔成功");

            System.Data.DataTable T2 = GetFMBG();

            for (int i = 0; i <= T2.Rows.Count - 1; i++)
            {

                DataRow row = T2.Rows[i];
                string LINENUM = row["LINENUM"].ToString();
                decimal DEBIT = Convert.ToDecimal(row["DEBIT"].ToString());
                AddAUOGD2F(LINENUM, DEBIT);
            }

        }

        private void account_FMD1DataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            
            e.Row.Cells["U_PC_BSTY1"].Value = "三聯式發票/電子計算機發票";
            e.Row.Cells["U_PC_BSINV"].Value = "__________";
            e.Row.Cells["U_PC_BSTY2"].Value = "應稅";
            e.Row.Cells["U_PC_BSTYI"].Value = "外加";
            e.Row.Cells["U_PC_BSCUS"].Value = "______________";
            e.Row.Cells["U_PC_BSNOT"].Value = "________";
            e.Row.Cells["U_PC_BSTY3"].Value = "500元以上";
            e.Row.Cells["U_PC_BSTY4"].Value = "費用";
            e.Row.Cells["U_PC_BSTYC"].Value = "一般";
            e.Row.Cells["U_PC_BSTAX"].Value = "0";
            e.Row.Cells["U_PC_BSAMN"].Value = "0";
            e.Row.Cells["U_PC_BSAMT"].Value = "0";
            e.Row.Cells["BATCHNUM"].Value = textBox1.Text;
            e.Row.Cells["TRANSID"].Value = textBox2.Text;
            e.Row.Cells["LineId"].Value = account_FMD1DataGridView.Rows.Count;
            e.Row.Cells["VisOrder"].Value = account_FMD1DataGridView.Rows.Count - 1;
            e.Row.Cells["U_PC_BSAPP"].Value = DateTime.Now.ToString("yyyyMMdd");
            e.Row.Cells["U_PC_BSDAT"].Value = DateTime.Now.ToString("yyyyMMdd");

            
        }
        public static bool IsNumber(string strNumber)
        {
            System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(@"^\d+(\.)?\d*$");
            return r.IsMatch(strNumber);
        }
        private void button3_Click(object sender, EventArgs e)
        {
            System.Data.DataTable H1 = GeOJDT();
            System.Data.DataTable H2 = GetFMB();
            System.Data.DataTable H3 = GetFMBSUM();
            System.Data.DataTable H4 = Get3SUM();
            System.Data.DataTable H5 = GetSUMOJDT();
            string H3T = H3.Rows[0][0].ToString();
            string H4T = H4.Rows[0][0].ToString();
            string H5T = H5.Rows[0][0].ToString();
            if (H3T != H4T)
            {
                MessageBox.Show("發票稅額跟SAP傳票憑證稅額不符合");
                return;
            }
            if (H3T != H5T)
            {
                MessageBox.Show("發票稅額跟SAP傳票稅額不符合");
                return;
            }
            if (IsNumber(textBox3.Text) == false)
            {
                MessageBox.Show("請輸入數字");
                return;
            }
            if (textBox3.Text == "")
            {
                MessageBox.Show("請輸入傳票號碼");
                return;
            }

            if (H1.Rows.Count > 0)
            {
                MessageBox.Show("日記帳已有紀錄");
                return;
            }

            this.Validate();
            this.account_FMD1BindingSource.EndEdit();
            this.account_FMD1TableAdapter.Update(this.accBank.Account_FMD1);
      

            for (int i = 0; i <= H2.Rows.Count - 1; i++)
            {

                DataRow drw = H2.Rows[i];


                string D1 = drw["U_PC_BSAPP"].ToString();
                string D3 = drw["U_PC_BSDAT"].ToString();
                string D4 = drw["U_PC_BSTAX"].ToString();
                string D5 = drw["U_PC_BSAMN"].ToString();
                string D6 = drw["U_PC_BSAMT"].ToString();
                if (D1.Length != 8 || D3.Length != 8)
                {
                    MessageBox.Show("日期格式不符");
                    return;
                }
                if (String.IsNullOrEmpty(D4))
                {
                    MessageBox.Show("稅額不可空白");
                    return;
                }

                try
                {
                    decimal DD4 = Convert.ToDecimal(D4);
                    decimal DD5 = Convert.ToDecimal(D5);
                    decimal DD6 = Convert.ToDecimal(D6);
                }
                catch
                {
                    MessageBox.Show("金額必須為數字");
                    return;
                }
            }
           AddAUOGD(textBox3.Text);
            //F1234
            for (int i = 0; i <= account_FMD1DataGridView.Rows.Count - 2; i++)
            {

                DataGridViewRow row;

                row = account_FMD1DataGridView.Rows[i];
                string LineId = row.Cells["LineId"].Value.ToString();
                string VisOrder = row.Cells["VisOrder"].Value.ToString();
                decimal U_PC_BSAMN = Convert.ToDecimal(row.Cells["U_PC_BSAMN"].Value);
                decimal U_PC_BSAMT = Convert.ToDecimal(row.Cells["U_PC_BSAMT"].Value);
                string D1 = row.Cells["U_PC_BSAPP"].Value.ToString();
                string D2 = D1.Substring(0, 4) + '.' + D1.Substring(4, 2) + '.' + D1.Substring(6, 2);
                DateTime U_PC_BSAPP = Convert.ToDateTime(D2);
                string U_PC_BSCUS = row.Cells["U_PC_BSCUS"].Value.ToString();
                string D3 = row.Cells["U_PC_BSDAT"].Value.ToString();
                string D4 = D3.Substring(0, 4) + '.' + D3.Substring(4, 2) + '.' + D3.Substring(6, 2);
                DateTime U_PC_BSDAT = Convert.ToDateTime(D4);
                string U_PC_BSINV = row.Cells["U_PC_BSINV"].Value.ToString();
                string U_PC_BSNOT = row.Cells["U_PC_BSNOT"].Value.ToString();
                decimal U_PC_BSTAX = Convert.ToDecimal(row.Cells["U_PC_BSTAX"].Value);
                string U_PC_BSTY1 = row.Cells["U_PC_BSTY1"].Value.ToString();
                string U_PC_BSTY2 = row.Cells["U_PC_BSTY2"].Value.ToString();
                string U_PC_BSTY3 = row.Cells["U_PC_BSTY3"].Value.ToString();
                string U_PC_BSTY4 = row.Cells["U_PC_BSTY4"].Value.ToString();
                string U_PC_BSTYC = row.Cells["U_PC_BSTYC"].Value.ToString();
                string U_PC_BSTYI = row.Cells["U_PC_BSTYI"].Value.ToString();
           
                AddAUOGD2(LineId, VisOrder, U_PC_BSAMN, U_PC_BSAMT, U_PC_BSAPP, U_PC_BSCUS, U_PC_BSDAT, U_PC_BSINV, U_PC_BSNOT, U_PC_BSTAX, U_PC_BSTY1, U_PC_BSTY2, U_PC_BSTY3, U_PC_BSTY4, U_PC_BSTYC, U_PC_BSTYI);
            }

            AddAUOGD3();

            MessageBox.Show("轉值成功");
        }
        public void AddAUOGD(string U_BSREN)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("Insert into DBO.[@CADMEN_FMD](DocEntry,DocNum,Period,Series,Handwrtten,Canceled,Object,UserSign,Transfered,Status,CreateDate,CreateTime,UpdateDate,UpdateTime,DataSource,U_BSREN) values(@DocEntry,@DocNum,@Period,@Series,@Handwrtten,@Canceled,@Object,@UserSign,@Transfered,@Status,@CreateDate,@CreateTime,@UpdateDate,@UpdateTime,@DataSource,@U_BSREN)", connection);
            command.CommandType = CommandType.Text;
            string T1 = GetONNM().Rows[0][0].ToString();
            string T2 = GetPEROID().Rows[0][0].ToString();
            command.Parameters.Add(new SqlParameter("@DocEntry",T1));
            command.Parameters.Add(new SqlParameter("@DocNum", T1));
            command.Parameters.Add(new SqlParameter("@Period", T2));
            command.Parameters.Add(new SqlParameter("@Series", 32));
            command.Parameters.Add(new SqlParameter("@Handwrtten", "N"));
            command.Parameters.Add(new SqlParameter("@Canceled", "N"));
            command.Parameters.Add(new SqlParameter("@Object", "FMD"));
            command.Parameters.Add(new SqlParameter("@UserSign", 4));
            command.Parameters.Add(new SqlParameter("@Transfered", "N"));
            command.Parameters.Add(new SqlParameter("@Status", "O"));
            command.Parameters.Add(new SqlParameter("@CreateDate", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@CreateTime", DateTime.Now.ToString("HHmm")));
            command.Parameters.Add(new SqlParameter("@UpdateDate",DateTime.Now));
            command.Parameters.Add(new SqlParameter("@UpdateTime",DateTime.Now.ToString("HHmm")));
            command.Parameters.Add(new SqlParameter("@DataSource", "I"));
            command.Parameters.Add(new SqlParameter("@U_BSREN", U_BSREN));
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

        public void AddOBTD(decimal LocTotal, int UserSign)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("Insert into DBO.[OBTD](BatchNum,Status,NumOfTrans,DateID,LocTotal,FcTotal,SysTotal,UserSign) values(@BatchNum,@Status,@NumOfTrans,@DateID,@LocTotal,0,@SysTotal,@UserSign)", connection);
            command.CommandType = CommandType.Text;
            string T1 = util.GetONNM2().Rows[0][0].ToString();
 
            command.Parameters.Add(new SqlParameter("@BatchNum", T1));
            command.Parameters.Add(new SqlParameter("@Status", "O"));
            command.Parameters.Add(new SqlParameter("@NumOfTrans", 1));
            command.Parameters.Add(new SqlParameter("@DateID", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@LocTotal", LocTotal));

            command.Parameters.Add(new SqlParameter("@SysTotal", LocTotal));
            command.Parameters.Add(new SqlParameter("@UserSign", UserSign));

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


        private System.Data.DataTable GetFMDD()
        {

            SqlConnection connection = new SqlConnection(globals.shipConnectionString);

            StringBuilder sb = new StringBuilder();
            sb.Append("SELECT DOCENTRY FROM DBO.[@CADMEN_FMD] WHERE U_BSREN=@U_BSREN");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_BSREN", textBox4.Text));
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
        public void FMDDELETE()
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" DELETE DBO.[@CADMEN_FMD] WHERE U_BSREN=@U_BSREN ", connection);
            command.CommandType = CommandType.Text;
            string T1 = GetONNM().Rows[0][0].ToString();
            string T2 = GetPEROID().Rows[0][0].ToString();
            command.Parameters.Add(new SqlParameter("@U_BSREN", textBox4.Text));
   
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
        public void FMD1DELETE(string DOCENTRY)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" DELETE DBO.[@CADMEN_FMD1] WHERE DOCENTRY=@DOCENTRY  ", connection);
            command.CommandType = CommandType.Text;
            string T1 = GetONNM().Rows[0][0].ToString();
            string T2 = GetPEROID().Rows[0][0].ToString();
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));

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
        //G1234
        public void UPDATEFMD(string U_PC_BSTY1, string U_PC_BSINV, string U_PC_BSTY2, string U_PC_BSTYI, string U_PC_BSCUS, string U_PC_BSAMN, string U_PC_BSTAX, string U_PC_BSAMT, string U_PC_BSNOT, string U_PC_BSDAT, string U_PC_BSAPP, string U_PC_BSTY3, string U_PC_BSTY4, string U_PC_BSTYC, string DOCENTRY, string LINEID)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE DBO.[@CADMEN_FMD1] SET U_PC_BSTY1=@U_PC_BSTY1,U_PC_BSINV=@U_PC_BSINV,U_PC_BSTY2=@U_PC_BSTY2,U_PC_BSTYI=@U_PC_BSTYI,U_PC_BSCUS=@U_PC_BSCUS,U_PC_BSAMN=@U_PC_BSAMN,U_PC_BSTAX=@U_PC_BSTAX,U_PC_BSAMT=@U_PC_BSAMT,U_PC_BSNOT=@U_PC_BSNOT,U_PC_BSDAT=@U_PC_BSDAT,U_PC_BSAPP=@U_PC_BSAPP,U_PC_BSTY3=@U_PC_BSTY3,U_PC_BSTY4=@U_PC_BSTY4,U_PC_BSTYC=@U_PC_BSTYC WHERE DOCENTRY=@DOCENTRY AND LINEID=@LINEID", connection);
            command.CommandType = CommandType.Text;
            string U_PC_BSTY1T = "";

//            三聯式發票/電子計算機發票
//三聯式收銀機發票
//二聯式收銀機/載有稅額之其他憑證
//海關代徵營業稅
//免用統一發票/收據
//三聯式、電子計算機、三聯式收銀機統一發票及一般稅額計算之電子發票之進貨退出或折讓證明單
//二聯式收銀機統一發票及載有稅額之其他憑證之進貨退出或折讓證明單
//進項海關退還溢繳營業稅申報單
//一般稅額計算之電子發票
            if (U_PC_BSTY1 == "三聯式發票/電子計算機發票")
            {
                U_PC_BSTY1T = "0";
            }
            else if (U_PC_BSTY1 == "三聯式收銀機發票")
            {
                U_PC_BSTY1T = "1";
            }
            else if (U_PC_BSTY1 == "二聯式收銀機/載有稅額之其他憑證")
            {
                U_PC_BSTY1T = "2";
            }
            else if (U_PC_BSTY1 == "海關代徵營業稅")
            {
                U_PC_BSTY1T = "3";
            }
            else if (U_PC_BSTY1 == "免用統一發票/收據")
            {
                U_PC_BSTY1T = "4";
            }
            else if (U_PC_BSTY1 == "三聯式、電子計算機、三聯式收銀機統一發票及一般稅額計算之電子發票之進貨退出或折讓證明單")
            {
                U_PC_BSTY1T = "5";
            }
            else if (U_PC_BSTY1 == "二聯式收銀機統一發票及載有稅額之其他憑證之進貨退出或折讓證明單")
            {
                U_PC_BSTY1T = "6";
            }
            else if (U_PC_BSTY1 == "進項海關退還溢繳營業稅申報單")
            {
                U_PC_BSTY1T = "7";
            }
            else if (U_PC_BSTY1 == "一般稅額計算之電子發票")
            {
                U_PC_BSTY1T = "8";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY1", U_PC_BSTY1T));
            command.Parameters.Add(new SqlParameter("@U_PC_BSINV", U_PC_BSINV));

            string U_PC_BSTY2T = "";
            if (U_PC_BSTY2 == "應稅")
            {
                U_PC_BSTY2T = "0";
            }
            else if (U_PC_BSTY2 == "免稅率")
            {
                U_PC_BSTY2T = "1";
            }
            else if (U_PC_BSTY2 == "零稅")
            {
                U_PC_BSTY2T = "2";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY2", U_PC_BSTY2T));
            string U_PC_BSTYIT = "";
            if (U_PC_BSTYI == "外加")
            {
                U_PC_BSTYIT = "0";
            }
            else if (U_PC_BSTYI == "內含")
            {
                U_PC_BSTYIT = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTYI", U_PC_BSTYIT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSCUS", U_PC_BSCUS));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAMN", U_PC_BSAMN));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTAX", U_PC_BSTAX));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAMT", U_PC_BSAMT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSNOT", U_PC_BSNOT));

            string D3 = U_PC_BSDAT;
            string D4 = D3.Substring(0, 4) + '.' + D3.Substring(4, 2) + '.' + D3.Substring(6, 2);
            command.Parameters.Add(new SqlParameter("@U_PC_BSDAT", D4));

            string D1 = U_PC_BSAPP;
            string D2 = D1.Substring(0, 4) + '.' + D1.Substring(4, 2) + '.' + D1.Substring(6, 2);
            command.Parameters.Add(new SqlParameter("@U_PC_BSAPP", D2));

            string U_PC_BSTY3T = "";
            if (U_PC_BSTY3 == "500元以上")
            {
                U_PC_BSTY3T = "0";
            }
            else if (U_PC_BSTY3 == "500元以下(含)")
            {
                U_PC_BSTY3T = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY3", U_PC_BSTY3T));
            string U_PC_BSTY4T = "";
            if (U_PC_BSTY4 == "進貨")
            {
                U_PC_BSTY4T = "0";
            }
            else if (U_PC_BSTY4 == "費用")
            {
                U_PC_BSTY4T = "1";
            }
            else if (U_PC_BSTY4 == "固定資產")
            {
                U_PC_BSTY4T = "2";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY4", U_PC_BSTY4T));
            string U_PC_BSTYCT = "";
            if (U_PC_BSTYC == "一般")
            {
                U_PC_BSTYCT = "0";
            }
            else if (U_PC_BSTYC == "作廢")
            {
                U_PC_BSTYCT = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTYC", U_PC_BSTYCT));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINEID", LINEID));
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

        public void UPDATEFMDIN(string LINEID,string VISORDER, string U_PC_BSAMN, string U_PC_BSAMT, string U_PC_BSAPP, string U_PC_BSCUS, string U_PC_BSDAT, string U_PC_BSINV, string U_PC_BSNOT, string U_PC_BSTAX, string U_PC_BSTY1, string U_PC_BSTY2, string U_PC_BSTY3, string U_PC_BSTY4, string U_PC_BSTYC, string U_PC_BSTYI)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("Insert into DBO.[@CADMEN_FMD1]([DocEntry],[LineId],[VisOrder],[Object],[U_PC_BSAMN],[U_PC_BSAMT],[U_PC_BSAPP],[U_PC_BSCUS],[U_PC_BSDAT],[U_PC_BSINV] ,[U_PC_BSNOT],[U_PC_BSRN1],[U_PC_BSTAX],[U_PC_BSTY1],[U_PC_BSTY2],[U_PC_BSTY3],[U_PC_BSTY4],[U_PC_BSTY5],[U_PC_BSTYC],[U_PC_BSTYI]) values(@DocEntry,@LineId,@VisOrder,@Object,@U_PC_BSAMN,@U_PC_BSAMT,@U_PC_BSAPP,@U_PC_BSCUS,@U_PC_BSDAT,@U_PC_BSINV,@U_PC_BSNOT,@U_PC_BSRN1,@U_PC_BSTAX,@U_PC_BSTY1,@U_PC_BSTY2,@U_PC_BSTY3,@U_PC_BSTY4,@U_PC_BSTY5,@U_PC_BSTYC,@U_PC_BSTYI)", connection);

            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", GetONNM().Rows[0][0].ToString()));
            command.Parameters.Add(new SqlParameter("@Object", "FMD"));
            string U_PC_BSTY1T = "";
            if (U_PC_BSTY1 == "三聯式發票/電子計算機發票")
            {
                U_PC_BSTY1T = "0";
            }
            else if (U_PC_BSTY1 == "三聯式收銀機發票")
            {
                U_PC_BSTY1T = "1";
            }
            else if (U_PC_BSTY1 == "二聯式收銀機/載有稅額之其他憑證")
            {
                U_PC_BSTY1T = "2";
            }
            else if (U_PC_BSTY1 == "海關代徵營業稅")
            {
                U_PC_BSTY1T = "3";
            }
            else if (U_PC_BSTY1 == "免用統一發票/收據")
            {
                U_PC_BSTY1T = "4";
            }
            else if (U_PC_BSTY1 == "三聯式、電子計算機、三聯式收銀機統一發票及一般稅額計算之電子發票之進貨退出或折讓證明單")
            {
                U_PC_BSTY1T = "5";
            }
            else if (U_PC_BSTY1 == "二聯式收銀機統一發票及載有稅額之其他憑證之進貨退出或折讓證明單")
            {
                U_PC_BSTY1T = "6";
            }
            else if (U_PC_BSTY1 == "進項海關退還溢繳營業稅申報單")
            {
                U_PC_BSTY1T = "7";
            }
            else if (U_PC_BSTY1 == "一般稅額計算之電子發票")
            {
                U_PC_BSTY1T = "8";
            }

            command.Parameters.Add(new SqlParameter("@U_PC_BSTY1", U_PC_BSTY1T));
            command.Parameters.Add(new SqlParameter("@U_PC_BSINV", U_PC_BSINV));

            string U_PC_BSTY2T = "";
            if (U_PC_BSTY2 == "應稅")
            {
                U_PC_BSTY2T = "0";
            }
            else if (U_PC_BSTY2 == "免稅率")
            {
                U_PC_BSTY2T = "1";
            }
            else if (U_PC_BSTY2 == "零稅")
            {
                U_PC_BSTY2T = "2";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY2", U_PC_BSTY2T));
            string U_PC_BSTYIT = "";
            if (U_PC_BSTYI == "外加")
            {
                U_PC_BSTYIT = "0";
            }
            else if (U_PC_BSTYI == "內含")
            {
                U_PC_BSTYIT = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTYI", U_PC_BSTYIT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSCUS", U_PC_BSCUS));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAMN", U_PC_BSAMN));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTAX", U_PC_BSTAX));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAMT", U_PC_BSAMT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSNOT", U_PC_BSNOT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSRN1", "________"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY5", "0"));
            string D3 = U_PC_BSDAT;
            string D4 = D3.Substring(0, 4) + '.' + D3.Substring(4, 2) + '.' + D3.Substring(6, 2);
            command.Parameters.Add(new SqlParameter("@U_PC_BSDAT", D4));

            string D1 = U_PC_BSAPP;
            string D2 = D1.Substring(0, 4) + '.' + D1.Substring(4, 2) + '.' + D1.Substring(6, 2);
            command.Parameters.Add(new SqlParameter("@U_PC_BSAPP", D2));

            string U_PC_BSTY3T = "";
            if (U_PC_BSTY3 == "500元以上")
            {
                U_PC_BSTY3T = "0";
            }
            else if (U_PC_BSTY3 == "500元以下(含)")
            {
                U_PC_BSTY3T = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY3", U_PC_BSTY3T));
            string U_PC_BSTY4T = "";
            if (U_PC_BSTY4 == "進貨")
            {
                U_PC_BSTY4T = "0";
            }
            else if (U_PC_BSTY4 == "費用")
            {
                U_PC_BSTY4T = "1";
            }
            else if (U_PC_BSTY4 == "固定資產")
            {
                U_PC_BSTY4T = "2";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY4", U_PC_BSTY4T));
            string U_PC_BSTYCT = "";
            if (U_PC_BSTYC == "一般")
            {
                U_PC_BSTYCT = "0";
            }
            else if (U_PC_BSTYC == "作廢")
            {
                U_PC_BSTYCT = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTYC", U_PC_BSTYCT));

            command.Parameters.Add(new SqlParameter("@LineId", LINEID));
            command.Parameters.Add(new SqlParameter("@VisOrder", VISORDER));
            
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
        public void AddAUOGD2(string LineId, string VisOrder, decimal U_PC_BSAMN, decimal U_PC_BSAMT, DateTime U_PC_BSAPP, string U_PC_BSCUS, DateTime U_PC_BSDAT, string U_PC_BSINV, string U_PC_BSNOT, decimal U_PC_BSTAX, string U_PC_BSTY1, string U_PC_BSTY2, string U_PC_BSTY3, string U_PC_BSTY4, string U_PC_BSTYC, string U_PC_BSTYI)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("Insert into DBO.[@CADMEN_FMD1]([DocEntry],[LineId],[VisOrder],[Object],[U_PC_BSAMN],[U_PC_BSAMT],[U_PC_BSAPP],[U_PC_BSCUS],[U_PC_BSDAT],[U_PC_BSINV] ,[U_PC_BSNOT],[U_PC_BSRN1],[U_PC_BSTAX],[U_PC_BSTY1],[U_PC_BSTY2],[U_PC_BSTY3],[U_PC_BSTY4],[U_PC_BSTY5],[U_PC_BSTYC],[U_PC_BSTYI]) values(@DocEntry,@LineId,@VisOrder,@Object,@U_PC_BSAMN,@U_PC_BSAMT,@U_PC_BSAPP,@U_PC_BSCUS,@U_PC_BSDAT,@U_PC_BSINV,@U_PC_BSNOT,@U_PC_BSRN1,@U_PC_BSTAX,@U_PC_BSTY1,@U_PC_BSTY2,@U_PC_BSTY3,@U_PC_BSTY4,@U_PC_BSTY5,@U_PC_BSTYC,@U_PC_BSTYI)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", GetONNM().Rows[0][0].ToString()));
            command.Parameters.Add(new SqlParameter("@LineId", LineId));
            command.Parameters.Add(new SqlParameter("@VisOrder", VisOrder));
            command.Parameters.Add(new SqlParameter("@Object", "FMD"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAMN", U_PC_BSAMN));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAMT", U_PC_BSAMT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAPP", U_PC_BSAPP));
            command.Parameters.Add(new SqlParameter("@U_PC_BSCUS", U_PC_BSCUS));
            command.Parameters.Add(new SqlParameter("@U_PC_BSDAT", U_PC_BSDAT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSINV", U_PC_BSINV));
            command.Parameters.Add(new SqlParameter("@U_PC_BSNOT", U_PC_BSNOT));
            command.Parameters.Add(new SqlParameter("@U_PC_BSRN1", "________"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTAX", U_PC_BSTAX));

            string U_PC_BSTY1T = "";
            if (U_PC_BSTY1 == "三聯式發票/電子計算機發票")
            {
                U_PC_BSTY1T = "0";
            }
            else if (U_PC_BSTY1 == "三聯式收銀機發票")
            {
                U_PC_BSTY1T = "1";
            }
            else if (U_PC_BSTY1 == "二聯式收銀機/載有稅額之其他憑證")
            {
                U_PC_BSTY1T = "2";
            }
            else if (U_PC_BSTY1 == "海關代徵營業稅")
            {
                U_PC_BSTY1T = "3";
            }
            else if (U_PC_BSTY1 == "免用統一發票/收據")
            {
                U_PC_BSTY1T = "4";
            }
            else if (U_PC_BSTY1 == "三聯式、電子計算機、三聯式收銀機統一發票及一般稅額計算之電子發票之進貨退出或折讓證明單")
            {
                U_PC_BSTY1T = "5";
            }
            else if (U_PC_BSTY1 == "二聯式收銀機統一發票及載有稅額之其他憑證之進貨退出或折讓證明單")
            {
                U_PC_BSTY1T = "6";
            }
            else if (U_PC_BSTY1 == "進項海關退還溢繳營業稅申報單")
            {
                U_PC_BSTY1T = "7";
            }
            else if (U_PC_BSTY1 == "一般稅額計算之電子發票")
            {
                U_PC_BSTY1T = "8";
            }


            command.Parameters.Add(new SqlParameter("@U_PC_BSTY1", U_PC_BSTY1T));
            string U_PC_BSTY2T = "";
            if (U_PC_BSTY2 == "應稅")
            {
                U_PC_BSTY2T = "0";
            }
            else if (U_PC_BSTY2 == "免稅率")
            {
                U_PC_BSTY2T = "1";
            }
            else if (U_PC_BSTY2 == "零稅")
            {
                U_PC_BSTY2T = "2";
            }

            command.Parameters.Add(new SqlParameter("@U_PC_BSTY2", U_PC_BSTY2T));

           string U_PC_BSTY3T = "";
            if (U_PC_BSTY3 == "500元以上")
            {
                U_PC_BSTY3T = "0";
            }
            else if (U_PC_BSTY3 == "500元以下(含)")
            {
                U_PC_BSTY3T = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY3", U_PC_BSTY3T));

            string U_PC_BSTY4T = "";
            if (U_PC_BSTY4 == "進貨")
            {
                U_PC_BSTY4T = "0";
            }
            else if (U_PC_BSTY4 == "費用")
            {
                U_PC_BSTY4T = "1";
            }
            else if (U_PC_BSTY4 == "固定資產")
            {
                U_PC_BSTY4T = "2";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY4", U_PC_BSTY4T));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY5", "0"));

            string U_PC_BSTYCT = "";
            if (U_PC_BSTYC == "一般")
            {
                U_PC_BSTYCT = "0";
            }
            else if (U_PC_BSTYC == "作廢")
            {
                U_PC_BSTYCT = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTYC", U_PC_BSTYCT));

            string U_PC_BSTYIT = "";
            if (U_PC_BSTYI == "外加")
            {
                U_PC_BSTYIT = "0";
            }
            else if (U_PC_BSTYI == "內含")
            {
                U_PC_BSTYIT = "1";
            }
            command.Parameters.Add(new SqlParameter("@U_PC_BSTYI", U_PC_BSTYIT));
            //U_PC_BSTYI
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
        public void AddAUOGD2F( string LINENUM, decimal U_PC_BSTAX)
        {



            SqlConnection connection = new SqlConnection(strCn);
            SqlCommand command = new SqlCommand("Insert into [Account_FMD1]([LineId],[VisOrder],[Object],[U_PC_BSAMN],[U_PC_BSAMT],[U_PC_BSAPP],[U_PC_BSCUS],[U_PC_BSDAT],[U_PC_BSINV] ,[U_PC_BSNOT],[U_PC_BSRN1],[U_PC_BSTAX],[U_PC_BSTY1],[U_PC_BSTY2],[U_PC_BSTY3],[U_PC_BSTY4],[U_PC_BSTY5],[U_PC_BSTYC],[U_PC_BSTYI],[BATCHNUM],[TRANSID],[LINENUM]) values(@LineId,@VisOrder,@Object,@U_PC_BSAMN,@U_PC_BSAMT,@U_PC_BSAPP,@U_PC_BSCUS,@U_PC_BSDAT,@U_PC_BSINV,@U_PC_BSNOT,@U_PC_BSRN1,@U_PC_BSTAX,@U_PC_BSTY1,@U_PC_BSTY2,@U_PC_BSTY3,@U_PC_BSTY4,@U_PC_BSTY5,@U_PC_BSTYC,@U_PC_BSTYI,@BATCHNUM,@TRANSID,@LINENUM)", connection);
            command.CommandType = CommandType.Text;
            string LINEID = GETFMDMAX().Rows[0]["LINEID"].ToString();
            string VISORDER = Convert.ToString(Convert.ToInt16(LINEID) - 1); 
            command.Parameters.Add(new SqlParameter("@LineId", LINEID));
            command.Parameters.Add(new SqlParameter("@VisOrder", VISORDER));
            command.Parameters.Add(new SqlParameter("@Object", "FMD"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAMN", "0"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAMT", "0"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSAPP", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@U_PC_BSCUS",  "______________"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSDAT", GetMenu.Day()));
            command.Parameters.Add(new SqlParameter("@U_PC_BSINV", "__________"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSNOT", "________"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSRN1", "________"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTAX", U_PC_BSTAX));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY1", "三聯式發票/電子計算機發票"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY2",  "應稅"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY3", "500元以上"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY4", "費用"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTY5", "0"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTYC", "一般"));
            command.Parameters.Add(new SqlParameter("@U_PC_BSTYI", "外加"));
            command.Parameters.Add(new SqlParameter("@BATCHNUM", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@TRANSID", textBox2.Text));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            //LINENUM
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
        public void AddAUOGD3()
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand(" UPDATE ONNM SET AUTOKEY=AUTOKEY+1 WHERE OBJECTCODE='FMD' ", connection);
            command.CommandType = CommandType.Text;

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
     
        private void account_FMD1DataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                if (account_FMD1DataGridView.Columns[e.ColumnIndex].Name == "U_PC_BSAMN" ||
                account_FMD1DataGridView.Columns[e.ColumnIndex].Name == "U_PC_BSTAX")
                {
                    decimal IQTY = 0;
                    decimal ITAX = 0;

                    IQTY = Convert.ToDecimal(this.account_FMD1DataGridView.Rows[e.RowIndex].Cells["U_PC_BSAMN"].Value);
                    ITAX = Convert.ToDecimal(this.account_FMD1DataGridView.Rows[e.RowIndex].Cells["U_PC_BSTAX"].Value);
                    this.account_FMD1DataGridView.Rows[e.RowIndex].Cells["U_PC_BSAMT"].Value = (IQTY + ITAX).ToString();

                }

                if (account_FMD1DataGridView.Columns[e.ColumnIndex].Name == "U_PC_BSAPP")
                {
                    string H1 = this.account_FMD1DataGridView.Rows[e.RowIndex].Cells["U_PC_BSAPP"].Value.ToString();
                    string H2 = "";
                    int t1 = H1.IndexOf(".");
                    if (t1 != -1)
                    {
                        H2 = H1.Substring(0, t1);
                        if (H2.Length == 1)
                        {
                            H2 = "0" + H2;
                        }


                        string H3 = H1.Substring(t1 + 1, H1.Length - t1 - 1);
                        if (H3.Length == 1)
                        {
                            H3 = "0" + H3;
                        }
                        this.account_FMD1DataGridView.Rows[e.RowIndex].Cells["U_PC_BSAPP"].Value = DateTime.Now.ToString("yyyy") + H2 + H3;
                    }
                }




                if (account_FMD1DataGridView.Columns[e.ColumnIndex].Name == "U_PC_BSDAT")
                {
                    string H1 = this.account_FMD1DataGridView.Rows[e.RowIndex].Cells["U_PC_BSDAT"].Value.ToString();
                    string H2 = "";
                    int t1 = H1.IndexOf(".");
                    if (t1 != -1)
                    {
                        H2 = H1.Substring(0, t1);
                        if (H2.Length == 1)
                        {
                            H2 = "0" + H2;
                        }


                        string H3 = H1.Substring(t1 + 1, H1.Length - t1 - 1);
                        if (H3.Length == 1)
                        {
                            H3 = "0" + H3;
                        }
                        this.account_FMD1DataGridView.Rows[e.RowIndex].Cells["U_PC_BSDAT"].Value = DateTime.Now.ToString("yyyy") + H2 + H3;
                    }

                    //U_PC_BSDAT
                }
            }
            catch { }
            //U_PC_BSAPP
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (IsNumber(textBox4.Text) == false)
            {
                MessageBox.Show("請輸入數字");
                return;
            }
            if (textBox4.Text == "")
            {
                MessageBox.Show("請輸入傳票號碼");
                return;
            }

            
            System.Data.DataTable J1 = GetFMD();

            dataGridView2.DataSource = J1;

            if (J1.Rows.Count == 0)
            {
                MessageBox.Show("SAP沒有資料,請您直接在此新增");
            }
            //if (J1.Rows.Count > 0)
            //{
            //    dataGridView2.DataSource = GetFMD();

            //}
            //else
            //{
            //  
            //}
        }

        private System.Data.DataTable GetFMD()
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("               SELECT CASE U_PC_BSTY1  ");
            sb.Append("               WHEN 0 THEN '三聯式發票/電子計算機發票' ");
            sb.Append("               WHEN 1 THEN '三聯式收銀機發票' ");
            sb.Append("               WHEN 2 THEN '二聯式收銀機/載有稅額之其他憑證' ");
            sb.Append("               WHEN 3 THEN '海關代徵營業稅' ");
            sb.Append("               WHEN 4 THEN '免用統一發票/收據' ");
            sb.Append(" 			                WHEN 5 THEN '三聯式、電子計算機、三聯式收銀機統一發票及一般稅額計算之電子發票之進貨退出或折讓證明單' ");
            sb.Append("               WHEN 6 THEN '二聯式收銀機統一發票及載有稅額之其他憑證之進貨退出或折讓證明單' ");
            sb.Append("               WHEN 7 THEN '進項海關退還溢繳營業稅申報單' ");
            sb.Append("               WHEN 8 THEN '一般稅額計算之電子發票' ");
            sb.Append("                END  U_PC_BSTY1,U_PC_BSINV,CASE U_PC_BSTY2 ");
            sb.Append(" WHEN 0 THEN '應稅' WHEN 1 THEN '免稅率' WHEN 2 THEN '零稅'");
            sb.Append("  END  U_PC_BSTY2,CASE U_PC_BSTYI");
            sb.Append(" WHEN 0 THEN '外加' WHEN 1 THEN '內含'");
            sb.Append("  END  U_PC_BSTYI,U_PC_BSCUS,U_PC_BSAMN,U_PC_BSTAX,U_PC_BSAMT,U_PC_BSNOT, CONVERT(VARCHAR(8),U_PC_BSDAT,112) U_PC_BSDAT, CONVERT(VARCHAR(8),U_PC_BSAPP,112) U_PC_BSAPP ");
            sb.Append(" ,U_PC_BSAPP,CASE U_PC_BSTY3");
            sb.Append(" WHEN 0 THEN '500元以上' WHEN 1 THEN '500元以下(含)' END  U_PC_BSTY3");
            sb.Append(" ,CASE U_PC_BSTY4");
            sb.Append(" WHEN 0 THEN '進貨' WHEN 1 THEN '費用' WHEN 2 THEN '固定資產'");
            sb.Append("  END  U_PC_BSTY4,CASE U_PC_BSTYC");
            sb.Append(" WHEN 0 THEN '一般' WHEN 1 THEN '作廢' ");
            sb.Append("  END  U_PC_BSTYC,T1.DOCENTRY,T1.LINEID,T1.VisOrder");
            sb.Append("  FROM dbo.[@CADMEN_FMD] T0");
            sb.Append(" INNER JOIN dbo.[@CADMEN_FMD1]  T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE T0.U_BSREN=@U_BSREN");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_BSREN", textBox4.Text));
            
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

        private void button5_Click(object sender, EventArgs e)
        {
            if (IsNumber(textBox4.Text) == false)
            {
                MessageBox.Show("請輸入數字");
                return;
            }
            if (textBox4.Text == "")
            {
                MessageBox.Show("請輸入傳票號碼");
                return;
            }

            System.Data.DataTable M1 = GetFMDD();
            if (M1.Rows.Count > 0)
            {
                FMDDELETE();
                FMD1DELETE(M1.Rows[0][0].ToString());


            }


            for (int i = 0; i <= dataGridView2.Rows.Count - 2; i++)
            {

                DataGridViewRow row;

                row = dataGridView2.Rows[i];
                string D1 = row.Cells["AU_PC_BSAPP"].Value.ToString();
                string D3 = row.Cells["AU_PC_BSDAT"].Value.ToString();
                string D4 = row.Cells["AU_PC_BSTAX"].Value.ToString();
                string D5 = row.Cells["AU_PC_BSAMN"].Value.ToString();
                string D6 = row.Cells["AU_PC_BSAMT"].Value.ToString();
                string AU_PC_BSINV = row.Cells["AU_PC_BSINV"].Value.ToString();
                
                if (D1.Length != 8 || D3.Length != 8)
                {
                    MessageBox.Show("日期格式不符");
                    return;
                }
                if (String.IsNullOrEmpty(D4))
                {
                    MessageBox.Show("稅額不可空白");
                    return;
                }
       
                try
                {
                    decimal DD4 = Convert.ToDecimal(D4);
                    decimal DD5 = Convert.ToDecimal(D5);
                    decimal DD6 = Convert.ToDecimal(D6);
                }
                catch
                {
                    MessageBox.Show("金額必須為數字");
                    return;
                }
            }


            AddAUOGD(textBox4.Text);
            for (int i = 0; i <= dataGridView2.Rows.Count - 2; i++)
            {

                DataGridViewRow row;

                row = dataGridView2.Rows[i];
                string U_PC_BSTY1 = row.Cells["AU_PC_BSTY1"].Value.ToString();
                string U_PC_BSINV = row.Cells["AU_PC_BSINV"].Value.ToString();
                string U_PC_BSTY2 = row.Cells["AU_PC_BSTY2"].Value.ToString();
                string U_PC_BSTYI = row.Cells["AU_PC_BSTYI"].Value.ToString();
                string U_PC_BSCUS = row.Cells["AU_PC_BSCUS"].Value.ToString();
                string U_PC_BSAMN = row.Cells["AU_PC_BSAMN"].Value.ToString();
                string U_PC_BSTAX = row.Cells["AU_PC_BSTAX"].Value.ToString();
                string U_PC_BSAMT = row.Cells["AU_PC_BSAMT"].Value.ToString();
                string U_PC_BSNOT = row.Cells["AU_PC_BSNOT"].Value.ToString();
                string U_PC_BSDAT = row.Cells["AU_PC_BSDAT"].Value.ToString();
                string U_PC_BSAPP = row.Cells["AU_PC_BSAPP"].Value.ToString();
                string U_PC_BSTY3 = row.Cells["AU_PC_BSTY3"].Value.ToString();
                string U_PC_BSTY4 = row.Cells["AU_PC_BSTY4"].Value.ToString();
                string U_PC_BSTYC = row.Cells["AU_PC_BSTYC"].Value.ToString();

                string LINEID = (i + 1).ToString();
                string VISORDER = i.ToString();
             //   AddAUOGD2(LINEID, LINEID, U_PC_BSAMN,U_PC_BSTY1, U_PC_BSINV, U_PC_BSTY2, U_PC_BSTYI, U_PC_BSCUS, U_PC_BSTAX, U_PC_BSAMT, U_PC_BSNOT, U_PC_BSDAT, U_PC_BSAPP, U_PC_BSTY3, U_PC_BSTY4, U_PC_BSTYC);
                if (U_PC_BSINV.Length == 10)
                {
                    UPDATEFMDIN(LINEID, VISORDER, U_PC_BSAMN, U_PC_BSAMT, U_PC_BSAPP, U_PC_BSCUS, U_PC_BSDAT, U_PC_BSINV, U_PC_BSNOT, U_PC_BSTAX, U_PC_BSTY1, U_PC_BSTY2, U_PC_BSTY3, U_PC_BSTY4, U_PC_BSTYC, U_PC_BSTYI);
                }
            }
            AddAUOGD3();
            MessageBox.Show("更新成功");
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

            try
            {
                if (dataGridView2.Columns[e.ColumnIndex].Name == "AU_PC_BSAMN" ||
                dataGridView2.Columns[e.ColumnIndex].Name == "AU_PC_BSTAX")
                {
                    decimal IQTY = 0;
                    decimal ITAX = 0;

                    IQTY = Convert.ToDecimal(this.dataGridView2.Rows[e.RowIndex].Cells["AU_PC_BSAMN"].Value);
                    ITAX = Convert.ToDecimal(this.dataGridView2.Rows[e.RowIndex].Cells["AU_PC_BSTAX"].Value);
                    this.dataGridView2.Rows[e.RowIndex].Cells["AU_PC_BSAMT"].Value = (IQTY + ITAX).ToString();

                }

                if (dataGridView2.Columns[e.ColumnIndex].Name == "AU_PC_BSAPP")
                {
                    string H1 = this.dataGridView2.Rows[e.RowIndex].Cells["AU_PC_BSAPP"].Value.ToString();
                    string H2 = "";
                    int t1 = H1.IndexOf(".");
                    if (t1 != -1)
                    {
                        H2 = H1.Substring(0, t1);
                        if (H2.Length == 1)
                        {
                            H2 = "0" + H2;
                        }


                        string H3 = H1.Substring(t1 + 1, H1.Length - t1 - 1);
                        if (H3.Length == 1)
                        {
                            H3 = "0" + H3;
                        }
                        this.dataGridView2.Rows[e.RowIndex].Cells["AU_PC_BSAPP"].Value = DateTime.Now.ToString("yyyy") + H2 + H3;
                    }
                }




                if (dataGridView2.Columns[e.ColumnIndex].Name == "AU_PC_BSDAT")
                {
                    string H1 = this.dataGridView2.Rows[e.RowIndex].Cells["AU_PC_BSDAT"].Value.ToString();
                    string H2 = "";
                    int t1 = H1.IndexOf(".");
                    if (t1 != -1)
                    {
                        H2 = H1.Substring(0, t1);
                        if (H2.Length == 1)
                        {
                            H2 = "0" + H2;
                        }


                        string H3 = H1.Substring(t1 + 1, H1.Length - t1 - 1);
                        if (H3.Length == 1)
                        {
                            H3 = "0" + H3;
                        }
                        this.dataGridView2.Rows[e.RowIndex].Cells["AU_PC_BSDAT"].Value = DateTime.Now.ToString("yyyy") + H2 + H3;
                    }

                    //U_PC_BSDAT
                }
            }
            catch { }
        }

        private void dataGridView2_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["AU_PC_BSTY1"].Value = "三聯式發票/電子計算機發票";
            e.Row.Cells["AU_PC_BSINV"].Value = "__________";
            e.Row.Cells["AU_PC_BSTY2"].Value = "應稅";
            e.Row.Cells["AU_PC_BSTYI"].Value = "外加";
            e.Row.Cells["AU_PC_BSCUS"].Value = "______________";
            e.Row.Cells["AU_PC_BSNOT"].Value = "________";
            e.Row.Cells["AU_PC_BSTY3"].Value = "500元以上";
            e.Row.Cells["AU_PC_BSTY4"].Value = "費用";
            e.Row.Cells["AU_PC_BSTYC"].Value = "一般";
            e.Row.Cells["AU_PC_BSTAX"].Value = "0";
            e.Row.Cells["AU_PC_BSAMN"].Value = "0";
            e.Row.Cells["AU_PC_BSAMT"].Value = "0";


            e.Row.Cells["ALINEID"].Value = dataGridView2.Rows.Count;
            e.Row.Cells["AAVISORDER"].Value = dataGridView2.Rows.Count - 1;
            e.Row.Cells["AU_PC_BSAPP"].Value = DateTime.Now.ToString("yyyyMMdd");
            e.Row.Cells["AU_PC_BSDAT"].Value = DateTime.Now.ToString("yyyyMMdd");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (textBox5.Text == "" )
            {
                MessageBox.Show("請輸入日記帳憑證號碼");
                return;

            }
            if (IsNumber(textBox5.Text) == false)
            {
                MessageBox.Show("日記帳憑證號碼請輸入數字");
                return;
            }

      

            System.Data.DataTable T1 = GetCOPY(textBox5.Text);
            System.Data.DataTable T12 = GetCOPY2(textBox5.Text);
            if (T1.Rows.Count == 0)
            {
                MessageBox.Show("分錄沒有資料");
                return;
            }


            dataGridView3.DataSource = T1;

            dataGridView4.DataSource = T12;
        }

        private void button7_Click(object sender, EventArgs e)
        {

               DialogResult result;
            result = MessageBox.Show("請確認是否要複製", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {
                if (textBox5.Text == "")
                {
                    MessageBox.Show("請輸入日記帳憑證號碼");
                    return;

                }
                if (IsNumber(textBox5.Text) == false)
                {
                    MessageBox.Show("日記帳憑證號碼請輸入數字");
                    return;
                }

                if (dataGridView3.Rows.Count == 0 || dataGridView4.Rows.Count == 0)
                {
                    MessageBox.Show("SAP沒有此單據資料");
                    return;

                }


                DataGridViewRow row = dataGridView3.Rows[0];

                decimal a0 = Convert.ToDecimal(row.Cells["總計"].Value);
                int USERSIGN = Convert.ToInt16(row.Cells["USERSIGN"].Value);
                string 製單人 = Convert.ToString(row.Cells["製單人"].Value);
                int FinncPriod = Convert.ToInt16(row.Cells["FinncPriod"].Value);
                util.AddOBTD(a0, USERSIGN);
                util.AddOBTF(製單人, a0, FinncPriod, USERSIGN,"");


                for (int h = 0; h <= dataGridView4.Rows.Count - 1; h++)
                {
                    DataGridViewRow row1 = dataGridView4.Rows[h];
                    int LINE = Convert.ToInt16(row1.Cells["LINE"].Value);
                    string 科目 = Convert.ToString(row1.Cells["科目"].Value);
                    decimal 借項 = Convert.ToDecimal(row1.Cells["借項"].Value);
                    decimal 貸項 = Convert.ToDecimal(row1.Cells["貸項"].Value);
                    string 摘要 = Convert.ToString(row1.Cells["摘要"].Value);
                    string TRANSTYPE = Convert.ToString(row1.Cells["TRANSTYPE"].Value);
                    string 專案 = Convert.ToString(row1.Cells["專案"].Value);
                    string 部門 = Convert.ToString(row1.Cells["部門"].Value);
                    string 稅群組 = Convert.ToString(row1.Cells["稅群組"].Value);
                    string VATLINE = Convert.ToString(row1.Cells["VATLINE"].Value);
                    string DEBCRED = Convert.ToString(row1.Cells["DEBCRED"].Value);
                    util.AddBTF1(LINE, 科目, 借項, 貸項, 摘要, TRANSTYPE, 製單人, 專案, 部門, USERSIGN, FinncPriod, 稅群組, VATLINE, DEBCRED);

                }
                string T1 = util.GetONNM2().Rows[0][0].ToString();
                MessageBox.Show("複製成功 新增號碼 : " + T1);
                util.ADDONNM();
            
            }
        }
    }
}

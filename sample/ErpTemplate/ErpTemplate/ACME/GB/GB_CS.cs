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
namespace ACME
{
    public partial class GB_CS : ACME.fmBase1
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHICOMP02;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public GB_CS()
        {
            InitializeComponent();
        }

        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            gB_CSTableAdapter.Connection = MyConnection;
            gB_CS1TableAdapter.Connection = MyConnection;
            gB_CS2TableAdapter.Connection = MyConnection;
            gB_CSDTableAdapter.Connection = MyConnection;
     
        }

        private void WW()
        {
            uSERNAMETextBox.ReadOnly = true;
            dOCDATETextBox.ReadOnly = true;
            rECAMTTextBox.ReadOnly = true;
            rECTOTALTextBox.ReadOnly = true;
        }
        public override void AfterCancelEdit()
        {
            WW();
        }
        public override void query()
        {
            sHIPPINGCODETextBox.ReadOnly = false;
            cARDCODETextBox.ReadOnly = false;
            cARDNAMETextBox.ReadOnly = false;

            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;

        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();

                pOTATO.GB_CS.RejectChanges();
                pOTATO.GB_CS1.RejectChanges();
                pOTATO.GB_CS2.RejectChanges();
                pOTATO.GB_CSD.RejectChanges();
            }
            catch
            {
            }
            return true;

        }
        public override void EndEdit()
        {
            WW();
        }
        public override void SetInit()
        {

            MyBS = gB_CSBindingSource;
            MyTableName = "GB_CS";
            MyIDFieldName = "SHIPPINGCODE";


        }
        public override void AfterEdit()
        {
            sHIPPINGCODETextBox.ReadOnly = true;
        }
        public override void AfterAddNew()
        {
            WW();

            dOCDATETextBox.Text = DateTime.Now.ToString("yyyyMMdd");

        }

        public override void SetDefaultValue()
        {
            if (kyes == null)
            {

                string NumberName = "CS" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                kyes = NumberName + AutoNum + "X";
            }
            this.sHIPPINGCODETextBox.Text = kyes;
            string username = fmLogin.LoginID.ToString();

            this.gB_CSBindingSource.EndEdit();
            kyes = null;
            dOCDATETextBox.Text = GetMenu.Day();
            sTATUSTextBox.Text = "受理";
            q73TTextBox.Text = "一般件";
            uSERNAMETextBox.Text = username;
            tabControl1.SelectedIndex = 1;

            q11CheckBox.Checked = false;
            q12CheckBox.Checked = false;
            q13CheckBox.Checked = false;
            q14CheckBox.Checked = false;
            q15CheckBox.Checked = false;
            q16CheckBox.Checked = false;
            q18CheckBox.Checked = false;
  
            q32CheckBox.Checked = false;
            q33CheckBox.Checked = false;
            q34CheckBox.Checked = false;
            q35CheckBox.Checked = false;
            q36CheckBox.Checked = false;
            q37CheckBox.Checked = false;
            q39CheckBox.Checked = false;
            q61CheckBox.Checked = false;
            q62CheckBox.Checked = false;
            q63CheckBox.Checked = false;
            q65CheckBox.Checked = false;
            q71CheckBox.Checked = false;
            q72CheckBox.Checked = false;

            q76CheckBox.Checked = false;
        
            tabControl1.SelectedIndex = 0;
            qTYPE1CheckBox.Checked = false;
            qTYPE2CheckBox.Checked = false;
            qTYPE3CheckBox.Checked = false;
            qTYPE4CheckBox.Checked = false;
            qTYPE5CheckBox.Checked = false;
            qTYPE6CheckBox.Checked = false;
        }
        public override void FillData()
        {
            try
            {
                gB_CSTableAdapter.Fill(pOTATO.GB_CS, MyID);
                gB_CS1TableAdapter.Fill(pOTATO.GB_CS1, MyID);
                gB_CS2TableAdapter.Fill(pOTATO.GB_CS2, MyID);
                gB_CSDTableAdapter.Fill(pOTATO.GB_CSD, MyID); 

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
                if (Convert.ToDouble(this.MyTableStatus) == 3)
                {

                    MyBS.RemoveAt(MyBS.Position);
                }

                Validate();

                gB_CS1BindingSource.MoveFirst();

                for (int i = 1; i <= gB_CS1BindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)gB_CS1BindingSource.Current;

                    row["SEQNO"] = i;



                    gB_CS1BindingSource.EndEdit();

                    gB_CS1BindingSource.MoveNext();
                }

                gB_CS2BindingSource.MoveFirst();

                for (int i = 1; i <= gB_CS2BindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)gB_CS2BindingSource.Current;

                    row["SEQNO"] = i;



                    gB_CS2BindingSource.EndEdit();

                    gB_CS2BindingSource.MoveNext();
                }

                gB_CSDBindingSource.MoveFirst();

                for (int i = 1; i <= gB_CSDBindingSource.Count; i++)
                {
                    DataRowView row = (DataRowView)gB_CSDBindingSource.Current;

                    row["SEQNO"] = i;

                    gB_CSDBindingSource.EndEdit();

                    gB_CSDBindingSource.MoveNext();
                }

                gB_CSTableAdapter.Connection.Open();


                gB_CSBindingSource.EndEdit();
                gB_CS1BindingSource.EndEdit();
                gB_CS2BindingSource.EndEdit();
                gB_CSDBindingSource.EndEdit();

                tx = gB_CSTableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(gB_CSTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter1 = util.GetAdapter(gB_CS1TableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter3 = util.GetAdapter(gB_CS2TableAdapter);
                Adapter3.UpdateCommand.Transaction = tx;
                Adapter3.InsertCommand.Transaction = tx;
                Adapter3.DeleteCommand.Transaction = tx;
                SqlDataAdapter Adapter2 = util.GetAdapter(gB_CSDTableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;


                gB_CSTableAdapter.Update(pOTATO.GB_CS);
                pOTATO.GB_CS.AcceptChanges();

                gB_CS1TableAdapter.Update(pOTATO.GB_CS1);
                pOTATO.GB_CS1.AcceptChanges();

                gB_CS2TableAdapter.Update(pOTATO.GB_CS2);
                pOTATO.GB_CS2.AcceptChanges();

                gB_CSDTableAdapter.Update(pOTATO.GB_CSD);
                pOTATO.GB_CSD.AcceptChanges();


                tx.Commit();

                this.MyID = this.sHIPPINGCODETextBox.Text;

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
                this.gB_CSTableAdapter.Connection.Close();

            }

 

            return UpdateData;
        }

        private void comboBox1_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBU("GBCS6");

            comboBox1.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            sTATUSTextBox.Text = comboBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUST();

            if (LookupValues != null)
            {
                string CARDCODE=Convert.ToString(LookupValues[0]);
                cARDCODETextBox.Text = CARDCODE;
                cARDNAMETextBox.Text = Convert.ToString(LookupValues[1]);

                System.Data.DataTable G1 = GETCUSTCON(CARDCODE);
                System.Data.DataTable G2 = GETCUSBU(CARDCODE);
                if (G1.Rows.Count > 0)
                {
                    rECTOTALTextBox.Text = G1.Rows[0][0].ToString();
                    rECAMTTextBox.Text = G1.Rows[0][1].ToString();

                }
                if (G2.Rows.Count > 0)
                {
                    dOCTYPETextBox.Text = G2.Rows[0][0].ToString();
                }
            }
        }
        public  System.Data.DataTable GETCUSBU(string NAME)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("            SELECT  L.EngName FROM comCustomer U     Left Join comCustClass L On U.ClassID =L.ClassID and L.Flag =1  WHERE ID = '" + NAME + "'");
         

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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
        public static System.Data.DataTable GETGBCSQ2(string QCODE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("            SELECT ID,QUESTION  FROM GB_CSQ2 where QCODE=@QCODE ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@QCODE", QCODE));
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
        public  System.Data.DataTable GETCHOITEM(string AA)
        {
            SqlConnection connection = new SqlConnection(strCn);
                        StringBuilder sb = new StringBuilder();
                        sb.Append("                SELECT * FROM (");
                        sb.Append("                SELECT  J.PRODID,J.PRODNAME ,   CASE ");
                        sb.Append("                                                                                     WHEN SUBSTRING(K.ClassID,3,1)='S' AND SUBSTRING(K.ClassID,1,1)='A'  THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'      ");
                        sb.Append("                                                                                     WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬'  WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '加工品' END 品項");
                        sb.Append("                                                                                      FROM comProduct J ");
                        sb.Append("                   Left Join comProductClass K On J.ClassID =K.ClassID   ");
                        sb.Append("              WHERE K.ClassID  <> 'ASC100'");
                        sb.Append("                      ) AS A WHERE ISNULL(品項,'') ='"+ AA +"' ");
                


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

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

        public  System.Data.DataTable GETBILLNO(string BillNO)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                     SELECT    A.BillNO 訂單編號,A.BillDate 訂購日期, A.LinkMan 到貨人,A.LinkTelephone 電話,A.CustAddress 地址,G.ProdID 料號,J.InvoProdName 名稱,G.Quantity  數量,G.RowNO   ");
            sb.Append("                FROM  OrdBillMain A  Inner Join OrdBillSub G  ");
            sb.Append("               On G.Flag=A.Flag  And G.BillNO=A.BillNO  ");
            sb.Append("               left join ComProdRec O On  O.FromNO=G.BillNO AND O.FromRow=G.RowNO  AND O.Flag =500  ");
            sb.Append("                   left join COMBILLACCOUNTS S ON (O.BillNO =S.FundBillNo AND S.Flag =500) ");
            sb.Append("                    left join comCustomer U On  U.ID=A.CustomerID AND U.Flag =1 ");
            sb.Append("                    Left Join comProduct B On B.ProdID=G.ProdID  ");
            sb.Append("             Left Join comProduct J On G.ProdID =J.ProdID   ");

            sb.Append("               WHERE A.Flag =2 AND   A.BillNO=@BillNO  ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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

        public System.Data.DataTable GETCUSTNAME(string BillNO)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();

            sb.Append("                       select DISTINCT   T0.CustomerID,U.FullName  from OrdBillMain T0  ");
            sb.Append("                                     left join comCustomer U On  U.ID=T0.CustomerID AND U.Flag =1 ");
            sb.Append("                              where  T0.BillNO =@BillNO  ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@BillNO", BillNO));
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
        public System.Data.DataTable GETCUSTCON(string CustomerID)
        {
            SqlConnection connection = new SqlConnection(strCn);
            StringBuilder sb = new StringBuilder();
            sb.Append("                select COUNT(BILLNO) COU, AVG(SumBTaxAmt) AMT  from OrdBillMain  WHERE CustomerID =@CustomerID ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@CustomerID", CustomerID));
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
        private void GB_CS_Load(object sender, EventArgs e)
        {

            WW();

            DataGridViewLinkColumn column = new DataGridViewLinkColumn();
            column.Name = "Link";
            column.UseColumnTextForLinkValue = true;
            column.Text = "讀取檔案";
            column.LinkBehavior = LinkBehavior.HoverUnderline;
            column.TrackVisitedState = true;
            gB_CSDDataGridView.Columns.Add(column);
        }

        private void comboBox5_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("GBCS1");

            comboBox5.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox5.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox5_SelectedIndexChanged(object sender, EventArgs e)
        {
            q21TextBox.Text = comboBox5.Text;
        }

        private void comboBox6_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("GBCS2");

            comboBox6.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox6.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox7_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("GBCS2");

            comboBox7.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox7.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox6_SelectedIndexChanged(object sender, EventArgs e)
        {
            q41TextBox.Text = comboBox6.Text;
        }

        private void comboBox7_SelectedIndexChanged(object sender, EventArgs e)
        {
            q51TextBox.Text = comboBox7.Text;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUSTGB2(q41TextBox.Text);

            if (LookupValues != null)
            {
                q42TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUSTGB2(q51TextBox.Text);

            if (LookupValues != null)
            {
                q52TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }


        private void comboBox2_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("GBCS5");

            comboBox2.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox2.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            dOCTYPETextBox.Text = comboBox2.Text;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUSTGB3(cARDCODETextBox.Text);

            if (LookupValues != null)
            {
               string BILLNO = Convert.ToString(LookupValues[0]);

               System.Data.DataTable dt1  = GETBILLNO(BILLNO);
               System.Data.DataTable dt2 = pOTATO.GB_CS1;
               for (int i = 0; i <= dt1.Rows.Count - 1; i++)
               {
                   DataRow drw = dt1.Rows[i];
                   DataRow drw2 = dt2.NewRow();
                   drw2["SHIPPINGCODE"] = sHIPPINGCODETextBox.Text;
                   drw2["BILLNO"] = drw["訂單編號"];
                   drw2["BILLDATE"] = drw["訂購日期"];
                   drw2["SHIPPERSON"] = drw["到貨人"];
                   drw2["SHIPTEL"] = drw["電話"];
                   drw2["SHIPADDRESS"] = drw["地址"];
                   drw2["ITEMCODE"] = drw["料號"];
                   drw2["ITEMNAME"] = drw["名稱"];
                   drw2["QTY"] = drw["數量"];
                   drw2["RowNO"] = drw["RowNO"];
                   dt2.Rows.Add(drw2);
               }
            }
        }

        private void comboBox3_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("GBCS2");

            comboBox3.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox3.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            q43TextBox.Text = comboBox3.Text;
        }

        private void comboBox8_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("GBCS2");

            comboBox8.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox8.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            q45TextBox.Text = comboBox8.Text;
        }

        private void comboBox9_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("GBCS2");

            comboBox9.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox9.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox9_SelectedIndexChanged(object sender, EventArgs e)
        {
            q53TextBox.Text = comboBox9.Text;
        }

        private void comboBox10_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("GBCS2");

            comboBox10.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox10.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox10_SelectedIndexChanged(object sender, EventArgs e)
        {
            q55TextBox.Text = comboBox10.Text;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUSTGB2(q43TextBox.Text);

            if (LookupValues != null)
            {
                q44TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUSTGB2(q45TextBox.Text);

            if (LookupValues != null)
            {
                q46TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUSTGB2(q53TextBox.Text);

            if (LookupValues != null)
            {
                q54TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUSTGB2(q55TextBox.Text);

            if (LookupValues != null)
            {
                q56TextBox.Text = Convert.ToString(LookupValues[1]);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                string server = "//acmesrv01//Public//ARMAS//客服附件//";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = download(filename);
                if (dt2.Rows.Count > 0)
                {
                    MessageBox.Show("檔案名稱重複,請修改檔名");
                }
                else
                {
                    if (result == DialogResult.OK)


                        MessageBox.Show(Path.GetFileName(opdf.FileName));
                    string file = opdf.FileName;
                    bool FF1 = getrma.UploadFile(file, server, false);
                    if (FF1 == false)
                    {
                        return;
                    }
                    System.Data.DataTable dt1 = pOTATO.GB_CSD;

                    DataRow drw = dt1.NewRow();
                    drw["SHIPPINGCODE"] = sHIPPINGCODETextBox.Text;

                    drw["filename"] = filename;
                    drw["path"] = @"\\acmesrv01\Public\ARMAS\客服附件\" + filename;
                    dt1.Rows.Add(drw);
                    this.gB_CSDBindingSource.EndEdit();
                    this.gB_CSDTableAdapter.Update(pOTATO.GB_CSD);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public static System.Data.DataTable download(string DocEntry)
        {
            SqlConnection MyConnection = globals.Connection;

            string sql = "select * from GB_CSD where [filename] = @DocEntry";
            SqlCommand command = new SqlCommand(sql, MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DocEntry", DocEntry));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                MyConnection.Open();
                da.Fill(ds, " AP_download ");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables[" AP_download "];
        }

        private void gB_CSDDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DataGridView dgv = (DataGridView)sender;


                if (dgv.Columns[e.ColumnIndex].Name == "Link")
                {
                    System.Data.DataTable dt1 = pOTATO.GB_CSD;
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

        private void comboBox11_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("G商品諮詢");

            comboBox11.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox11.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox11_SelectedIndexChanged(object sender, EventArgs e)
        {
            qTYPE11TextBox.Text = comboBox11.Text;
        }

        private void comboBox12_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("G訂購問題");

            comboBox12.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox12.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            qTYPE21TextBox.Text = comboBox12.Text;
        }

        private void comboBox13_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("G付款問題");

            comboBox13.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox13.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            qTYPE31TextBox.Text = comboBox13.Text;
        }

        private void comboBox14_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("G配送問題");

            comboBox14.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox14.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox14_SelectedIndexChanged(object sender, EventArgs e)
        {
            qTYPE41TextBox.Text = comboBox14.Text;
        }

        private void comboBox15_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("G退換貨問題");

            comboBox15.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox15.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox15_SelectedIndexChanged(object sender, EventArgs e)
        {
            qTYPE51TextBox.Text = comboBox15.Text;
        }

        private void comboBox16_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("G其他問題");

            comboBox16.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox16.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox16_SelectedIndexChanged(object sender, EventArgs e)
        {
            qTYPE61TextBox.Text = comboBox16.Text;
        }

        private void comboBox17_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("G案件性質");

            comboBox17.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox17.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            q73TTextBox.Text = comboBox17.Text;
        }

        private void comboBox18_MouseClick(object sender, MouseEventArgs e)
        {
            System.Data.DataTable dt3 = GetMenu.GetBUGB("G案件來源");

            comboBox18.Items.Clear();


            for (int i = 0; i <= dt3.Rows.Count - 1; i++)
            {
                comboBox18.Items.Add(Convert.ToString(dt3.Rows[i][0]));
            }
        }

        private void comboBox18_SelectedIndexChanged(object sender, EventArgs e)
        {
            q74TTextBox.Text = comboBox18.Text;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
            {
                MessageBox.Show("請輸入電話號碼");
                return;
            }
            object[] LookupValues = null;
            LookupValues = GetMenu.GetCHICUSTGB4(textBox1.Text.Replace("&", "").Replace("-", "").Replace("(", "").Replace(")", ""));

            if (LookupValues != null)
            {
                string BILLNO = Convert.ToString(LookupValues[0]);

                System.Data.DataTable dt1 = GETBILLNO(BILLNO);

                System.Data.DataTable dt2 = pOTATO.GB_CS1;

                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {

                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();
                    drw2["SHIPPINGCODE"] = sHIPPINGCODETextBox.Text;
                    drw2["BILLNO"] = drw["訂單編號"];
                    drw2["BILLDATE"] = drw["訂購日期"];
                    drw2["SHIPPERSON"] = drw["到貨人"];
                    drw2["SHIPTEL"] = drw["電話"];
                    drw2["SHIPADDRESS"] = drw["地址"];
                    drw2["ITEMCODE"] = drw["料號"];
                    drw2["ITEMNAME"] = drw["名稱"];
                    drw2["QTY"] = drw["數量"];
                    drw2["RowNO"] = drw["RowNO"];
                    dt2.Rows.Add(drw2);


                }





                for (int j = 0; j <= gB_CS1DataGridView.Rows.Count - 2; j++)
                {
                    gB_CS1DataGridView.Rows[j].Cells[0].Value = j.ToString();
                }

                System.Data.DataTable G3 = GETCUSTNAME(BILLNO);
                if (G3.Rows.Count > 0)
                {
                    cARDCODETextBox.Text = G3.Rows[0][0].ToString();
                    cARDNAMETextBox.Text = G3.Rows[0][0].ToString();


                }

                System.Data.DataTable G1 = GETCUSTCON(cARDCODETextBox.Text);
                System.Data.DataTable G2 = GETCUSBU(cARDCODETextBox.Text);
                if (G1.Rows.Count > 0)
                {
                    rECTOTALTextBox.Text = G1.Rows[0][0].ToString();
                    rECAMTTextBox.Text = G1.Rows[0][1].ToString();

                }
                if (G2.Rows.Count > 0)
                {
                    dOCTYPETextBox.Text = G2.Rows[0][0].ToString();
                }

                gB_CSBindingSource.EndEdit();
                gB_CS1BindingSource.EndEdit();
            }
        }

            
         
        

    

        //private void FormatColumn(string strconn,string AA)
        //{
        //    try
        //    {
        //        dataGridView2.Columns.Remove("ZIPCODE");
        //        DataGridViewComboBoxColumn combo = new DataGridViewComboBoxColumn();
        //        System.Data.DataTable dt = new System.Data.DataTable();
        //        using (SqlConnection conn = new SqlConnection(strconn))
        //        {
        //            StringBuilder sb = new StringBuilder();
        //            sb.Append("                SELECT * FROM (");
        //            sb.Append("                SELECT  J.PRODID,J.PRODNAME ,   CASE ");
        //            sb.Append("                                                                                     WHEN SUBSTRING(K.ClassID,3,1)='S' AND SUBSTRING(K.ClassID,1,1)='A'  THEN '蝦' WHEN SUBSTRING(K.ClassID,3,1)='C' THEN '雞'      ");
        //            sb.Append("                                                                                     WHEN SUBSTRING(K.ClassID,3,1)='P' THEN '豬'  WHEN SUBSTRING(K.ClassID,1,3)='BPK' THEN '加工品' END 品項");
        //            sb.Append("                                                                                      FROM comProduct J ");
        //            sb.Append("                   Left Join comProductClass K On J.ClassID =K.ClassID   ");
        //            sb.Append("              WHERE K.ClassID  <> 'ASC100'");
        //            sb.Append("                      ) AS A WHERE ISNULL(品項,'') ='"+ AA +"' ");
        //            using (SqlDataAdapter da = new SqlDataAdapter(sb.ToString(), conn))
        //            {
        //                try
        //                {
        //                    da.Fill(dt);

        //                }
        //                catch (SqlException sqlex)
        //                {
        //                    throw new Exception(sqlex.Message);
        //                }
        //                catch (Exception ex)
        //                {
        //                    throw new Exception(ex.Message);
        //                }
        //            }

        //        }
        //        combo.DisplayIndex = 2;
        //        combo.HeaderText = "ZIPCODE";
        //        combo.DataPropertyName = "PRODID";//資料行名稱
        //        combo.DisplayMember = "PRODNAME";//顯示清單選項內容
        //        combo.ValueMember = "PRODID"; //清單選項對應的值
        //        combo.DataSource = dt;

        //        dataGridView2.Columns.Insert(1, combo);
        //    }
        //    catch
        //    { }
        //}

  
    }
}

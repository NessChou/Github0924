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
    public partial class PROBOM : ACME.fmBase1
    {

        public PROBOM()
        {
            InitializeComponent();
        }
        public string PublicString2;
        public override void SetConnection()
        {
            MyConnection = globals.Connection;
            sOLAR_PROBOMTableAdapter.Connection = MyConnection;
            sOLAR_PROBOM2TableAdapter.Connection = MyConnection;
            sOLAR_PROBOMDownloadTableAdapter.Connection = MyConnection;
        }
        private void WW()
        {
            shippingCodeTextBox.ReadOnly = true;
            vERTextBox.ReadOnly = true;
            button3.Enabled = true;
            button1.Enabled = true;
            button2.Enabled = true;
            dOCDATETextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            button5.Enabled = true;
            button7.Enabled = true;
            button8.Enabled = true;
        }
        public override void query()
        {
            shippingCodeTextBox.ReadOnly = false;
            dOCDATETextBox.ReadOnly = false;
            createNameTextBox.ReadOnly = false;
        }
        public override bool BeforeCancelEdit()
        {
            try
            {
                Validate();

                sOLAR.SOLAR_PROBOM.RejectChanges();
                sOLAR.SOLAR_PROBOM2.RejectChanges();
                sOLAR.SOLAR_PROBOMDownload.RejectChanges();
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

            TOTAL();
        }
        public override void AfterEdit()
        {


            sOLAR_PROBOM2DataGridView.Columns["PROJECTCODE"].ReadOnly = true;
            sOLAR_PROBOM2DataGridView.Columns["PROJECTNAME"].ReadOnly = true;
            sOLAR_PROBOM2DataGridView.Columns["CHILDNUM"].ReadOnly = true;
            sOLAR_PROBOM2DataGridView.Columns["ITEMCODE"].ReadOnly = true;
            sOLAR_PROBOM2DataGridView.Columns["ITEMNAME"].ReadOnly = true;
            sOLAR_PROBOM2DataGridView.Columns["PCOST"].ReadOnly = true;
            sOLAR_PROBOM2DataGridView.Columns["OPCOST"].ReadOnly = true;
            sOLAR_PROBOM2DataGridView.Columns["FATHER"].ReadOnly = true;
            sOLAR_PROBOM2DataGridView.Columns["DOCTYPE"].ReadOnly = true;
            
            shippingCodeTextBox.ReadOnly = true;
            dOCDATETextBox.ReadOnly = true;
            createNameTextBox.ReadOnly = true;
            vERTextBox.ReadOnly = true;
           

        }
        public override void AfterAddNew()
        {
            WW();
        }

        public override bool BeforeEndEdit()
        {

            int J1 = 0;
            bool BeforeEndEdit;
            BeforeEndEdit = true;
            for (int i = 0; i <= sOLAR_PROBOM2DataGridView.Rows.Count - 2; i++)
            {
                
                DataGridViewRow row;

                row = sOLAR_PROBOM2DataGridView.Rows[i];
                string T1 = row.Cells["COST"].Value.ToString();
                if (String.IsNullOrEmpty(T1))
                {
                    T1 = "0";
                }
                decimal a0 = Convert.ToDecimal(T1);
                string a1 = row.Cells["ID"].Value.ToString();
                string DOCTYPE = row.Cells["DOCTYPE"].Value.ToString().Trim();
                string FATHER = row.Cells["FATHER"].Value.ToString().Trim();
                string ITEMCODE = row.Cells["ITEMCODE"].Value.ToString().Trim();

                decimal a2 = 0;
                System.Data.DataTable H1 = GetBOMLOG(a1);
                if (H1.Rows.Count > 0)
                {
                    a2 = Convert.ToDecimal(H1.Rows[0][0].ToString());
                }

                if (a0 != a2)
                {
                    J1 = 1;
                    AddBOM(a0, a2, DOCTYPE, vERTextBox.Text, FATHER, ITEMCODE);
                }
            }
            if (J1 == 1)
            {
                DialogResult result;
                result = MessageBox.Show("預估成本已修改過，請確定是否要更新版本", "YES/NO", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    string H1 = vERTextBox.Text;
                    if (H1 == "")
                    {
                        vERTextBox.Text = "0";
                    }
                    else
                    {
                        int T1 = Convert.ToInt16(H1) + 1;
                        vERTextBox.Text = T1.ToString();
                    }

                    sOLAR_PROBOMBindingSource.EndEdit();
                    sOLAR_PROBOMTableAdapter.Update(sOLAR.SOLAR_PROBOM);
                    sOLAR.SOLAR_PROBOM.AcceptChanges();

                    for (int i = 0; i <= sOLAR_PROBOM2DataGridView.Rows.Count - 2; i++)
                    {
                        //OPCOST 採購成本
                        //PCOST 已付採購成本
                        //PRECOST 未付採購成本
                        //COST 預估成本

                        DataGridViewRow row;
                        
                        row = sOLAR_PROBOM2DataGridView.Rows[i];
                      //  int CHILDNUM = Convert.ToInt16(row.Cells["CHILDNUM"].Value.ToString().Trim());
                        string PROJECTCODE = row.Cells["PROJECTCODE"].Value.ToString().Trim();
                        string PROJECTNAME = row.Cells["PROJECTNAME"].Value.ToString().Trim();
                        string DOCTYPE = row.Cells["DOCTYPE"].Value.ToString().Trim();
                        string FATHER = row.Cells["FATHER"].Value.ToString().Trim();
                        string ITEMCODE = row.Cells["ITEMCODE"].Value.ToString().Trim();
                        string ITEMNAME = row.Cells["ITEMNAME"].Value.ToString().Trim();
                        string COST1 = row.Cells["COST"].Value.ToString().Trim();
                        string PCOST1 = row.Cells["PCOST"].Value.ToString().Trim();
                        string PRECOST1 = row.Cells["PRECOST"].Value.ToString().Trim();
                        string OPCOST1 = row.Cells["OPCOST"].Value.ToString().Trim();
                        string QTY1 = row.Cells["QTY"].Value.ToString().Trim();
                        string PRICE1 = row.Cells["PRICE"].Value.ToString().Trim();
                        string DOCENTRY = row.Cells["DOCENTRY"].Value.ToString().Trim();
              
                        if (String.IsNullOrEmpty(COST1))
                        {
                            COST1 = "0";
                        }
                        if (String.IsNullOrEmpty(PCOST1))
                        {
                            PCOST1 = "0";
                        }
                        if (String.IsNullOrEmpty(PRECOST1))
                        {
                            PRECOST1 = "0";
                        }
                        if (String.IsNullOrEmpty(OPCOST1))
                        {
                            OPCOST1 = "0";
                        }
                        if (String.IsNullOrEmpty(QTY1))
                        {
                            QTY1 = "0";
                        }
                        if (String.IsNullOrEmpty(PRICE1))
                        {
                            PRICE1 = "0";
                        }
                        decimal COST = Convert.ToDecimal(COST1);
                        decimal PCOST = Convert.ToDecimal(PCOST1);
                        decimal PRECOST = Convert.ToDecimal(PRECOST1);
                        decimal OPCOST = Convert.ToDecimal(OPCOST1);
                        decimal QTY = Convert.ToDecimal(QTY1);
                        decimal PRICE = Convert.ToDecimal(PRICE1);
                        string VER = vERTextBox.Text;
                        string OWORDOC = row.Cells["OWORDOC"].Value.ToString().Trim();
                        AddBOM4(PROJECTCODE, PROJECTNAME, ITEMCODE, ITEMNAME, FATHER, QTY, PRICE, OPCOST, PCOST, PRECOST, COST, DOCTYPE, DOCENTRY, VER, OWORDOC);
                    }
                    MessageBox.Show("版本已更新");
                }
            }
            return BeforeEndEdit;
        }
        public override void SetInit()
        {

            MyBS = sOLAR_PROBOMBindingSource;
            MyTableName = "SOLAR_PROBOM";
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
            string username = fmLogin.LoginID.ToString();
            dOCDATETextBox.Text = GetMenu.Day();
            createNameTextBox.Text = username;
            this.sOLAR_PROBOMBindingSource.EndEdit();
            kyes = null;
        }
        public override void FillData()
        {
            try
            {

                if (!String.IsNullOrEmpty(PublicString2))
                {
                    MyID = PublicString2;

                }

                sOLAR_PROBOMTableAdapter.Fill(sOLAR.SOLAR_PROBOM, MyID);
                sOLAR_PROBOM2TableAdapter.Fill(sOLAR.SOLAR_PROBOM2, MyID);
                sOLAR_PROBOMDownloadTableAdapter.Fill(sOLAR.SOLAR_PROBOMDownload, MyID);
                TOTAL();
               
                   for (int i = 0; i <= sOLAR_PROBOM2DataGridView.Rows.Count - 2; i++)
                   {

                       DataGridViewRow row;

                       row = sOLAR_PROBOM2DataGridView.Rows[i];
                       string a0 = row.Cells["DOCTYPE"].Value.ToString().Trim();
                       string OWORDOC = row.Cells["OWORDOC"].Value.ToString().Trim();
                       string OWORLINE = row.Cells["OWORLINE"].Value.ToString().Trim();
                       string COST = row.Cells["COST"].Value.ToString().Trim();
                       if (!String.IsNullOrEmpty(COST))
                       {
                           if (COST != "0")
                           {
                               if (!String.IsNullOrEmpty(OWORDOC))
                               {
                                   System.Data.DataTable DTH = DT3(OWORDOC, OWORLINE);
                                   if (DTH.Rows.Count == 0)
                                   {
                                       UpdateOWOR(OWORDOC, OWORLINE);
                                   }
                               }
                           }
                       }

                       if (globals.GroupID.ToString().Trim() == "SOLAR_1")
                       {
                           if (a0 == "PV" || a0 == "INV")
                           {
                               if (i == 0)
                               {
                                   sOLAR_PROBOM2DataGridView.CurrentCell = sOLAR_PROBOM2DataGridView.Rows[1].Cells[0];
                               }
                               sOLAR_PROBOM2DataGridView.Rows[i].Visible = false;
                           }
      
                       }
            
                }
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

                sOLAR_PROBOM2BindingSource.MoveFirst();

                for (int i = 0; i <= sOLAR_PROBOM2BindingSource.Count - 1; i++)
                {
                    DataRowView row3 = (DataRowView)sOLAR_PROBOM2BindingSource.Current;

                    row3["CHILDNUM"] = i;



                    sOLAR_PROBOM2BindingSource.EndEdit();

                    sOLAR_PROBOM2BindingSource.MoveNext();

                }

                sOLAR_PROBOMTableAdapter.Connection.Open();


                sOLAR_PROBOMBindingSource.EndEdit();
                sOLAR_PROBOM2BindingSource.EndEdit();


                tx = sOLAR_PROBOMTableAdapter.Connection.BeginTransaction();


                SqlDataAdapter Adapter = util.GetAdapter(sOLAR_PROBOMTableAdapter);
                Adapter.UpdateCommand.Transaction = tx;
                Adapter.InsertCommand.Transaction = tx;
                Adapter.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter1 = util.GetAdapter(sOLAR_PROBOM2TableAdapter);
                Adapter1.UpdateCommand.Transaction = tx;
                Adapter1.InsertCommand.Transaction = tx;
                Adapter1.DeleteCommand.Transaction = tx;

                SqlDataAdapter Adapter2 = util.GetAdapter(sOLAR_PROBOMDownloadTableAdapter);
                Adapter2.UpdateCommand.Transaction = tx;
                Adapter2.InsertCommand.Transaction = tx;
                Adapter2.DeleteCommand.Transaction = tx;


                sOLAR_PROBOMTableAdapter.Update(sOLAR.SOLAR_PROBOM);
                sOLAR.SOLAR_PROBOM.AcceptChanges();

                sOLAR_PROBOM2TableAdapter.Update(sOLAR.SOLAR_PROBOM2);
                sOLAR.SOLAR_PROBOM2.AcceptChanges();

                sOLAR_PROBOMDownloadTableAdapter.Update(sOLAR.SOLAR_PROBOMDownload);
                sOLAR.SOLAR_PROBOMDownload.AcceptChanges();


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
                this.sOLAR_PROBOMTableAdapter.Connection.Close();

            }
            return UpdateData;
        }
        private System.Data.DataTable GetBOM(string U_PROJECTCODE)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT MAX(T0.ITEMCODE) 母件編號,T1.ITEMCODE 子件編號,MAX(T2.ITEMNAME) 產品名稱,");
            sb.Append("      T3.LINENUM LINENUM,");
            sb.Append("       MAX(U_GROUP)  DTYPE,MAX(T3.U_MEMO)  MEMO,MAX(T2.BUYUNITMSR) 單位,");
            sb.Append("      T3.DOCENTRY,MAX(T3.QUANTITY) QTY,MAX(T3.PRICE) PRICE,MAX(T3.LINETOTAL) TOTAL,MAX(T0.U_PROJECTCODE) 專案代碼,MAX(T1.DOCENTRY) OWORDOC,MAX(T1.LINENUM) OWORLINE ");
            sb.Append("        FROM OWOR T0");
            sb.Append("       LEFT JOIN WOR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("       LEFT JOIN OITM T2 ON (T1.ITEMCODE=T2.ITEMCODE)");
            sb.Append("       LEFT JOIN POR1 T3 ON (T0.U_PROJECTCODE=T3.PROJECT AND T1.ITEMCODE=T3.ITEMCODE)");
            sb.Append("        LEFT JOIN OPOR T4 ON(T3.DOCENTRY=T4.DOCENTRY) ");
            sb.Append(" WHERE T0.U_PROJECTCODE in (" + U_PROJECTCODE + ") AND ISNULL(T4.CANCELED,'') <> 'Y' AND  T0.STATUS<> 'C'    ");
            sb.Append(" GROUP BY T3.DOCENTRY,T1.ITEMCODE,T3.LINENUM     ORDER BY MAX(T1.DOCENTRY),MAX(T1.LINENUM) ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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


            return ds.Tables[0];

        }

        private System.Data.DataTable GetAMOUNT(string FATHER,string CODE)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(PRICE,0) PRICE,ISNULL(CAST(QUANTITY AS INT),0) QTY,U_DESC, ISNULL(PRICE,0)*ISNULL(CAST(QUANTITY AS INT),0) COST  FROM ITT1 WHERE FATHER=@FATHER AND CODE=@CODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@FATHER", FATHER));
            command.Parameters.Add(new SqlParameter("@CODE", CODE));

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


            return ds.Tables[0];

        }

        private System.Data.DataTable GetQTY(string FATHER, string ITEMCODE2)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT SUM(T1.PLANNEDQTY) 計劃數量,SUM(T1.ISSUEDQTY) 已發貨,SUM(T1.BASEQTY) 基礎數量 FROM OWOR T0");
            sb.Append(" LEFT JOIN WOR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append("  WHERE T0.ITEMCODE=@FATHER AND T1.ITEMCODE=@ITEMCODE2");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@FATHER", FATHER));
            command.Parameters.Add(new SqlParameter("@ITEMCODE2", ITEMCODE2));

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


            return ds.Tables[0];

        }
        private System.Data.DataTable GetBOM2(string DOCENTRY, string LINENUM)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  SUM(T1.LINETOTAL)  LINETOTAL FROM OPOR T0 ");
            sb.Append(" LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE    T1.DOCENTRY=@DOCENTRY AND  T1.LINENUM=@LINENUM AND LINESTATUS='C' HAVING ISNULL(SUM(T1.LINETOTAL),0)  <> 0 ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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


            return ds.Tables[0];

        }


        private System.Data.DataTable GetPRJNAME(string PRJCODE)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  PRJNAME FROM OPRJ WHERE  PRJCODE=@PRJCODE ");
          
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@PRJCODE", PRJCODE));

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


            return ds.Tables[0];

        }
        private System.Data.DataTable GetBOM4(string DOCENTRY, string LINENUM)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT  T1.LINETOTAL LINETOTAL,T1.QUANTITY,T1.PRICE FROM OPOR T0 ");
            sb.Append(" LEFT JOIN POR1 T1 ON (T0.DOCENTRY=T1.DOCENTRY)");
            sb.Append(" WHERE    T1.DOCENTRY=@DOCENTRY AND  T1.LINENUM=@LINENUM  ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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


            return ds.Tables[0];

        }



        private System.Data.DataTable GetBOM3(string ITEMCODE)
        {

            SqlConnection MyConnection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT AVGPRICE FROM OITM WHERE ITEMCODE=@ITEMCODE ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));

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


            return ds.Tables[0];

        }


        private void TOTAL() 
        {
          System.Data.DataTable  dtCost = MakeTableCombine();

          System.Data.DataTable dtt = GetT1();
          DataRow dr = null;
          for (int i = 0; i <= dtt.Rows.Count - 1; i++)
          {
              dr = dtCost.NewRow();
              dr["NO1"] = (sOLAR_PROBOM2DataGridView.Rows.Count + i).ToString();
              dr["專案代碼"] = "";
              dr["專案名稱"] = "";
              dr["類型"] = "";
              dr["母件編號"] = "";
              dr["子件編號"] = "";
              dr["產品名稱"] = dtt.Rows[i]["產品名稱"].ToString().Trim();
              dr["規格說明"] =  "";
              dr["數量"] = "";
              dr["單價"] = "";
              dr["採購單號"] = "";
              dr["採購成本"] = Convert.ToDecimal(dtt.Rows[i]["採購成本"]);
              dr["已付採購成本"] = Convert.ToDecimal(dtt.Rows[i]["已付採購成本"]);
              dr["預付採購金額"] = Convert.ToDecimal(dtt.Rows[i]["採購成本"]) - Convert.ToDecimal(dtt.Rows[i]["已付採購成本"]);
              dr["預估成本"] = "0";
              dr["需求日期"] = "";
              dr["其他說明"] = dtt.Rows[i]["其他說明"].ToString();
            
              
              dtCost.Rows.Add(dr);
          }


          System.Data.DataTable dt = sOLAR.SOLAR_PROBOM2;
          decimal[] Total = new decimal[dt.Columns.Count - 1];

          for (int i = 0; i <= dt.Rows.Count - 1; i++)
          {

              for (int j = 11; j <= 14; j++)
              {
                  if (!String.IsNullOrEmpty(dt.Rows[i][j].ToString()))
                  {
                      string t1 = dt.Rows[i][j].ToString();
                      Total[j - 1] += Convert.ToInt64(dt.Rows[i][j]);
                  }

              }
          }


          for (int i = 0; i <= dtCost.Rows.Count - 1; i++)
          {

              for (int j = 11; j <= 14; j++)
              {
                  if (!String.IsNullOrEmpty(dtCost.Rows[i][j].ToString()))
                  {
                      string t1 = dtCost.Rows[i][j].ToString();
                      Total[j - 1] += Convert.ToInt64(dtCost.Rows[i][j]);
                  }

              }
          }

          DataRow row;

          row = dtCost.NewRow();

          row[5] = "合計";
          for (int j = 11; j <= 14; j++)
          {
              row[j] = Total[j - 1];
          }


          dtCost.Rows.Add(row);

          for (int i = 11; i <= 14; i++)
          {
              DataGridViewColumn col = dataGridView1.Columns[i];


              col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

              col.DefaultCellStyle.Format = "#,##0";
          }

          dataGridView1.DataSource = dtCost;

        }
        private void button2_Click(object sender, EventArgs e)
        {
                       DialogResult result;
                       result = MessageBox.Show("請確定是否要更新版本", "YES/NO", MessageBoxButtons.YesNo);
            if (result == DialogResult.Yes)
            {

          
            }
        }

      


        private System.Data.DataTable GetBOMLOG(string ID)
        {

            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT ISNULL(COST,0) FROM SOLAR_PROBOM2 WHERE ID=@ID ");
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ID", ID));

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


            return ds.Tables[0];

        }

        private System.Data.DataTable GetT1()
        {
            string ff = pROJECTTextBox.Text;
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append(" SELECT oPITEM 產品名稱,OPAMT 採購成本,AMT 已付採購成本,SHIPPINGCODE 其他說明 FROM SOLAR_PAY  WHERE dOCTYPE='預付請款'  AND  ISNULL(PAYCHECK,'') <> 'True' ");
            if (ff == "")
            {
                sb.Append("  AND PRJID in ('') ");
            }
            else
            {
                sb.Append("  AND PRJID in (" + ff + ") ");
            }
            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
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


            return ds.Tables[0];

        }
        public void AddBOM(decimal NCOST, decimal OCOST, string DOCTYPE, string VER,string FATHER,string ITEMCODE)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into SOLAR_PROBOM3(ShippingCode,NCOST,OCOST,UPUSER,UDATE,DOCTYPE,VER,FATHER,ITEMCODE) values(@ShippingCode,@NCOST,@OCOST,@UPUSER,@UDATE,@DOCTYPE,@VER,@FATHER,@ITEMCODE)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@NCOST", NCOST));
            command.Parameters.Add(new SqlParameter("@OCOST", OCOST));
            command.Parameters.Add(new SqlParameter("@UPUSER", fmLogin.LoginID.ToString()));
            command.Parameters.Add(new SqlParameter("@UDATE", DateTime.Now));
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@FATHER", FATHER));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
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
        public void AddBOM4(string PROJECTCODE, string PROJECTNAME, string ITEMCODE, string ITEMNAME, string FATHER, decimal QTY, decimal PRICE, decimal OPCOST, decimal PCOST, decimal PRECOST, decimal COST, string DOCTYPE, string DOCENTRY, string VER, string OWORDOC)
        {
            SqlConnection connection = new SqlConnection(globals.ConnectionString);
            SqlCommand command = new SqlCommand("Insert into SOLAR_PROBOM4(ShippingCode,PROJECTCODE,PROJECTNAME,ITEMCODE,ITEMNAME,FATHER,QTY,PRICE,OPCOST,PCOST,PRECOST,COST,DOCTYPE,DOCENTRY,VER,DOCDATE,OWORDOC) values(@ShippingCode,@PROJECTCODE,@PROJECTNAME,@ITEMCODE,@ITEMNAME,@FATHER,@QTY,@PRICE,@OPCOST,@PCOST,@PRECOST,@COST,@DOCTYPE,@DOCENTRY,@VER,@DOCDATE,@OWORDOC)", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@ShippingCode", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@PROJECTCODE", PROJECTCODE));
            command.Parameters.Add(new SqlParameter("@PROJECTNAME", PROJECTNAME));
            command.Parameters.Add(new SqlParameter("@ITEMCODE", ITEMCODE));
            command.Parameters.Add(new SqlParameter("@ITEMNAME", ITEMNAME));
            command.Parameters.Add(new SqlParameter("@FATHER", FATHER));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));

            command.Parameters.Add(new SqlParameter("@OPCOST", OPCOST));
            command.Parameters.Add(new SqlParameter("@PCOST", PCOST));
            command.Parameters.Add(new SqlParameter("@PRECOST", PRECOST));
            command.Parameters.Add(new SqlParameter("@COST", COST));
            command.Parameters.Add(new SqlParameter("@DOCTYPE", DOCTYPE));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@VER", VER));
            command.Parameters.Add(new SqlParameter("@DOCDATE", DateTime.Now.ToString("yyyyMMdd")));
            command.Parameters.Add(new SqlParameter("@OWORDOC", OWORDOC));
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
            PROLOG frm1 = new PROLOG();
            frm1.N1 = shippingCodeTextBox.Text;
            if (frm1.ShowDialog() == DialogResult.OK)
            {

            }
        }
        public override void STOP()
        {
          
                //for (int i = 0; i <= sOLAR_PROBOM2DataGridView.Rows.Count - 2; i++)
                //{
                //    DataGridViewRow row;
                //    row = sOLAR_PROBOM2DataGridView.Rows[i];
                //    string OWORDOC = row.Cells["OWORDOC"].Value.ToString().Trim();
                //    string OWORLINE = row.Cells["OWORLINE"].Value.ToString().Trim();

                //    System.Data.DataTable DTH = DT4(OWORDOC, OWORLINE);
                //    if (DTH.Rows.Count > 0)
                //    {
                //        string SHIP = DTH.Rows[0][0].ToString();
                //        MessageBox.Show("專案 ID " + SHIP + " 已成立，無法再新增");
                //        this.SSTOPID = "1";
                //        return;
                //    }
                //}
            

        }
        private void PROBOM_Load(object sender, EventArgs e)
        {
            WW();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            object[] LookupValues = GetMenu.GetSOLAR();

            if (LookupValues != null)
            {
              
                //pRJNAMETextBox.Text = Convert.ToString(LookupValues[1]);
                StringBuilder sb = new StringBuilder();


                for (int i = 0; i <= LookupValues.Length - 1; i++)
                {

                    sb.Append("'" + Convert.ToString(LookupValues[i]) + "',");

                }
                sb.Remove(sb.Length - 1, 1);
                string ds = sb.ToString();
                pROJECTTextBox.Text = ds;


                System.Data.DataTable dt1 = GetBOM(ds);
                System.Data.DataTable dt2 = sOLAR.SOLAR_PROBOM2;
                string OWORLINE = "";
                for (int i = 0; i <= dt1.Rows.Count - 1; i++)
                {
                    //OPCOST 採購成本
                    //PCOST 已付採購成本
                    //PRECOST 未付採購成本
                    //COST 預估成本

                    DataRow drw = dt1.Rows[i];
                    DataRow drw2 = dt2.NewRow();

                    string PROJECTCODE = drw["專案代碼"].ToString();
                    drw2["ShippingCode"] = shippingCodeTextBox.Text;

                    drw2["OWORDOC"] = drw["OWORDOC"];
                    drw2["OWORLINE"] = drw["OWORLINE"];
                    
                    drw2["LINENUM"] = drw["LINENUM"];
                    drw2["FATHER"] = drw["母件編號"];
                    drw2["ITEMCODE"] = drw["子件編號"];
                
                    drw2["ITEMNAME"] = drw["產品名稱"];
                    drw2["U_MEMO"] = drw["MEMO"];
                    drw2["DOCENTRY"] = drw["DOCENTRY"];
                    drw2["DOCTYPE"] = drw["DTYPE"];
                    drw2["OPCOST"] = drw["TOTAL"];
              
                    drw2["QTY"] = drw["QTY"];
                    drw2["PRICE"] = drw["PRICE"];
                    drw2["PROJECTCODE"] = PROJECTCODE;
                    System.Data.DataTable dt3P = GetPRJNAME(PROJECTCODE);
                    drw2["PROJECTNAME"] = dt3P.Rows[0][0].ToString();
                    drw2["UNIT"] = drw["單位"];
                    string FATHER = drw["母件編號"].ToString();
                    string CHILD = drw["子件編號"].ToString();
                    string LINENUM = drw["LINENUM"].ToString();
                    string DOCENTRY = drw["DOCENTRY"].ToString();
             

                    System.Data.DataTable dt3 = GetBOM2(DOCENTRY, LINENUM);
                    System.Data.DataTable dt4T = GetAMOUNT(FATHER, CHILD);
                    System.Data.DataTable dt5T = GetQTY(FATHER, CHILD);
                    decimal OPCOST = 0;
                    decimal PCOST = 0;
                    if (dt3.Rows.Count > 0)
                    {
                        drw2["PCOST"] = dt3.Rows[0][0].ToString();
                        PCOST = Convert.ToDecimal(dt3.Rows[0][0].ToString());
                    }


                    if (dt4T.Rows.Count > 0)
                    {
                    
                        if (String.IsNullOrEmpty(drw2["QTY"].ToString()))
                        {
                            drw2["QTY"] = dt4T.Rows[0][1].ToString();
                        }

                        if (drw["OWORLINE"].ToString() == OWORLINE)
                        {
                            drw2["COST"] = 0;
                        }
                        else
                        {

                            drw2["COST"] = dt4T.Rows[0]["COST"].ToString();
                        }
                        //decimal A1 = Convert.ToDecimal(dt4T.Rows[0][0]);
                        //decimal A2 = Convert.ToDecimal(drw2["QTY"]);
                        //drw2["COST"] = Convert.ToDecimal(dt4T.Rows[0][0]) * Convert.ToDecimal(drw2["QTY"]); 
                    }
                    if (String.IsNullOrEmpty(drw2["QTY"].ToString()))
                    {
                        if (dt5T.Rows.Count > 0)
                        {
                            drw2["QTY"] = dt5T.Rows[0]["計劃數量"].ToString();
                        }
                    }
                    if (String.IsNullOrEmpty(drw2["PRICE"].ToString()))
                    {
                        System.Data.DataTable dt4 = GetBOM3(CHILD);
                        drw2["PRICE"] = dt4.Rows[0][0].ToString();

                    }
                    if (String.IsNullOrEmpty(drw2["OPCOST"].ToString()))
                    {
                        System.Data.DataTable dt4 = GetBOM3(CHILD);
                        drw2["OPCOST"] = dt4.Rows[0][0].ToString();

                        if (dt4.Rows.Count > 0 )
                        {
                            if (dt4T.Rows.Count > 0)
                            {
                                drw2["OPCOST"] = Convert.ToDecimal(dt4.Rows[0][0]) * Convert.ToDecimal(dt4T.Rows[0][1]);
                            }
                            else if (dt5T.Rows.Count > 0)
                            {
                                drw2["OPCOST"] = Convert.ToDecimal(dt4.Rows[0][0]) * Convert.ToDecimal(dt5T.Rows[0][0]);
                            }
                            drw2["PCOST"] = drw2["OPCOST"];
                            PCOST = Convert.ToDecimal(drw2["OPCOST"].ToString());
                        }
                       
                    }
                    OPCOST = Convert.ToDecimal(drw2["OPCOST"]);
                    drw2["PRECOST"] = OPCOST - PCOST;

                    //if (String.IsNullOrEmpty(drw2["PCOST"].ToString()))
                    //{
                    //    if (dt5T.Rows.Count > 0)
                    //    {
                           
                    //        drw2["PCOST"] = Convert.ToDecimal(dt5T.Rows[0]["已發貨"]) * Convert.ToDecimal(drw2["PRICE"]);
                    //    }
                    //}

                    //int SA = Convert.ToInt32(drw2["PRECOST"]);

                    //if (SA == 0)
                    //{
                    //    if (dt5T.Rows.Count > 0)
                    //    {

                    //        drw2["PRECOST"] = (Convert.ToDecimal(dt5T.Rows[0]["基礎數量"]) - Convert.ToDecimal(dt5T.Rows[0]["已發貨"])) * Convert.ToDecimal(drw2["PRICE"]);
                    //    }
                    //}
                    
                    for (int j = 0; j <= sOLAR_PROBOM2DataGridView.Rows.Count - 2; j++)
                    {
                        sOLAR_PROBOM2DataGridView.Rows[j].Cells[0].Value = j.ToString();
                    }
                     OWORLINE = drw["OWORLINE"].ToString();
                    dt2.Rows.Add(drw2);
                }

                sOLAR_PROBOMBindingSource.EndEdit();
                sOLAR_PROBOM2BindingSource.EndEdit();
            }
        }

        private System.Data.DataTable MakeTableCombine()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("NO1", typeof(string));
            dt.Columns.Add("專案代碼", typeof(string));
            dt.Columns.Add("專案名稱", typeof(string));
            dt.Columns.Add("類型", typeof(string));
            dt.Columns.Add("母件編號", typeof(string));
            dt.Columns.Add("子件編號", typeof(string));
            dt.Columns.Add("產品名稱", typeof(string));
            dt.Columns.Add("規格說明", typeof(string));
            dt.Columns.Add("採購單號", typeof(string));
            dt.Columns.Add("數量", typeof(string));
            dt.Columns.Add("單價", typeof(string));
            dt.Columns.Add("採購成本", typeof(decimal));
            dt.Columns.Add("已付採購成本", typeof(decimal));
            dt.Columns.Add("預付採購金額", typeof(decimal));
            dt.Columns.Add("預估成本", typeof(decimal));
            dt.Columns.Add("需求日期", typeof(string));
            dt.Columns.Add("其他說明", typeof(string));
            return dt;
        }
        private System.Data.DataTable MakeTableCombine2()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("專案代碼", typeof(string));
            dt.Columns.Add("專案名稱", typeof(string));
            dt.Columns.Add("子件編號", typeof(string));
            dt.Columns.Add("產品名稱", typeof(string));
            dt.Columns.Add("採購成本", typeof(string));
            dt.Columns.Add("已付採購成本", typeof(string));
            dt.Columns.Add("未付採購成本", typeof(string));
            dt.Columns.Add("期初預估成本", typeof(string));
            dt.Columns.Add("預估未付", typeof(string));
            dt.Columns.Add("實際未來總成本", typeof(string));
            return dt;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            PROVER frm1 = new PROVER();
            frm1.N2 = shippingCodeTextBox.Text;
            if (frm1.ShowDialog() == DialogResult.OK)
            {

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            System.Data.DataTable DT1 = DT();

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\SOLAR\\請購確認單.xls";
            
            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            ExcelReport.ExcelReportOutput(DT1, ExcelTemplate, OutPutFile, "N");
        }
        private System.Data.DataTable DT()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT PROJECT 專案代碼,PRJNAME 專案名稱,CREATENAME 製表人,DOCDATE 製表日期,VER 版本");
            sb.Append(" ,CHILDNUM [NO],DOCTYPE 類別,''''+ITEMCODE 子料號,ITEMNAME 產品名稱,QTY 數量");
            sb.Append(" ,COST 預估金額,REQDATE 需求日期,MEMO 其他說明,DESCRIP 規格說明 FROM SOLAR_PROBOM T0");
            sb.Append(" LEFT JOIN SOLAR_PROBOM2 T1 ON (T0.SHIPPINGCODE=T1.SHIPPINGCODE)");
            sb.Append(" WHERE T0.SHIPPINGCODE=@SHIPPINGCODE ORDER BY CAST(CHILDNUM AS INT) ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text.ToString()));
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

        private System.Data.DataTable DT2()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("     SELECT PROJECTCODE 專案代碼,PROJECTNAME 專案名稱,''''+ITEMCODE 子件編號,ITEMNAME 產品名稱");
            sb.Append("         ,SUM(OPCOST) 採購成本,ISNULL(SUM(PCOST),0)  已付採購成本,ISNULL(SUM(PRECOST),0) 未付採購成本,ISNULL(SUM(COST),0)  期初預估成本");
            sb.Append("         FROM  SOLAR_PROBOM2");
            sb.Append(" WHERE SHIPPINGCODE=@SHIPPINGCODE ");
            sb.Append(" GROUP BY PROJECTCODE ,PROJECTNAME ,ITEMCODE ,ITEMNAME ");
            sb.Append("  ORDER BY PROJECTCODE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text.ToString()));
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
        private System.Data.DataTable DT21()
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append("              SELECT CHILDNUM [NO],PROJECTCODE 專案代碼,PROJECTNAME 專案名稱,DOCTYPE 類型,''''+ITEMCODE 子件編號,ITEMNAME 產品名稱 ");
            sb.Append(" ,DOCENTRY 採購單號,QTY 數量,UNIT 單位,PRICE 單價");
            sb.Append("                       ,ISNULL(OPCOST,0) 採購成本,ISNULL(PCOST,0)  已付採購成本,ISNULL(PRECOST,0) 未付採購成本,ISNULL(COST,0)  預估成本,U_MEMO 備註,LINENUM");
            sb.Append("                       FROM  SOLAR_PROBOM2 WHERE SHIPPINGCODE=@SHIPPINGCODE");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text.ToString()));
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
        private System.Data.DataTable DT3(string DOCENTRY, string LINENUM)
        {

            SqlConnection connection = globals.shipConnection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT * FROM WOR1 WHERE DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
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

        private System.Data.DataTable DT4(string OWORDOC, string OWORLINE)
        {

            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();

            sb.Append("  SELECT SHIPPINGCODE FROM SOLAR_PROBOM2 WHERE OWORDOC=@OWORDOC AND OWORLINE=@OWORLINE ");


            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@OWORDOC", OWORDOC));
            command.Parameters.Add(new SqlParameter("@OWORLINE", OWORLINE));
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
        private void button25_Click(object sender, EventArgs e)
        {

            string[] filebType = Directory.GetDirectories("//ACMEW08R2AP//SAPFILES//AttachmentsSolar2001//");
            string dd = DateTime.Now.ToString("yyyyMM");

            try
            {
                string server = "//ACMEW08R2AP//SAPFILES//AttachmentsSolar2001///";
                OpenFileDialog opdf = new OpenFileDialog();
                DialogResult result = opdf.ShowDialog();
                string filename = Path.GetFileName(opdf.FileName);
                System.Data.DataTable dt2 = GetMenu.download2(filename);

                if (dt2.Rows.Count > 0)
                {
                    MessageBox.Show("檔案名稱重複,請修改檔名");
                }
                else
                {
                    if (result == DialogResult.OK)
                    {

                        string file = opdf.FileName;
                        bool FF1 = getrma.UploadFile(file, server, false);
                        if (FF1 == false)
                        {
                            return;
                        }
                        System.Data.DataTable dt1 = sOLAR.SOLAR_PROBOMDownload;

                        DataRow drw = dt1.NewRow();
                        drw["ShippingCode"] = shippingCodeTextBox.Text;
                        drw["seq"] = (sOLAR_PROBOMDownloadDataGridView.Rows.Count).ToString();
                        drw["filename"] = filename;
                        string de = DateTime.Now.ToString("yyyyMM") + "\\";
                        drw["path"] = @"\\ACMEW08R2AP\SAPFILES\AttachmentsSolar2001\" + filename;
                        dt1.Rows.Add(drw);

                       sOLAR_PROBOMDownloadBindingSource.MoveFirst();

                       for (int i = 0; i <= sOLAR_PROBOMDownloadBindingSource.Count - 1; i++)
                        {
                            DataRowView rowd = (DataRowView)sOLAR_PROBOMDownloadBindingSource.Current;

                            rowd["seq"] = i;



                            sOLAR_PROBOMDownloadBindingSource.EndEdit();

                            sOLAR_PROBOMDownloadBindingSource.MoveNext();
                        }

                       this.sOLAR_PROBOMDownloadBindingSource.EndEdit();
                        this.sOLAR_PROBOMDownloadTableAdapter.Update(sOLAR.SOLAR_PROBOMDownload);
                        sOLAR.SOLAR_PROBOMDownload.AcceptChanges();

                        MessageBox.Show("上傳成功");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void sOLAR_PROBOMDownloadDataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                DataGridView dgv = (DataGridView)sender;

                if (dgv.Columns[e.ColumnIndex].Name == "LINK")
                {

                    System.Data.DataTable dt1 = sOLAR.SOLAR_PROBOMDownload;
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

        private void button6_Click(object sender, EventArgs e)
        {
            for (int j = 0; j <= sOLAR_PROBOM2DataGridView.Rows.Count - 1; j++)
            {
                //COST
                string LINENUM = sOLAR_PROBOM2DataGridView.Rows[j].Cells["LINENUM"].Value.ToString();
                string DOCENTRY = sOLAR_PROBOM2DataGridView.Rows[j].Cells["DOCENTRY"].Value.ToString();
                string CHILD = sOLAR_PROBOM2DataGridView.Rows[j].Cells["ITEMCODE"].Value.ToString();
                string FATHER = sOLAR_PROBOM2DataGridView.Rows[j].Cells["FATHER"].Value.ToString();

                string COST = sOLAR_PROBOM2DataGridView.Rows[j].Cells["COST"].Value.ToString();
                System.Data.DataTable dt3 = GetBOM2(DOCENTRY, LINENUM);
                System.Data.DataTable dt3T = GetBOM4(DOCENTRY, LINENUM);


                string PCOST = "0";
                string OPCOST = "0";
                string QUANTITY = "0";
                string PRICE1 = "0";
                if (dt3.Rows.Count > 0)
                {
                    PCOST = dt3.Rows[0][0].ToString();
                }

                if (dt3T.Rows.Count > 0)
                {
                    OPCOST = dt3T.Rows[0][0].ToString();
                    QUANTITY = dt3T.Rows[0]["QUANTITY"].ToString();
                    PRICE1 = dt3T.Rows[0]["PRICE"].ToString();
                }



                if (COST == "0.00" || String.IsNullOrEmpty(COST))
                {
                    if (dt3.Rows.Count > 0 || dt3T.Rows.Count > 0)
                    {
                        UpdateCOST(OPCOST, PCOST, LINENUM, DOCENTRY, Convert.ToDecimal(OPCOST) - Convert.ToDecimal(PCOST), Convert.ToDecimal(QUANTITY), Convert.ToDecimal(PRICE1));
                    }
                }
            }

            sOLAR_PROBOM2TableAdapter.Fill(sOLAR.SOLAR_PROBOM2, MyID);
        }

        private void UpdateOWOR(string OWORDOC, string OWORLINE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE  SOLAR_PROBOM2 SET COST=0,MEMO='生產訂單已刪除' WHERE OWORDOC=@OWORDOC AND OWORLINE=@OWORLINE ");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@OWORDOC", OWORDOC));
            command.Parameters.Add(new SqlParameter("@OWORLINE", OWORLINE));
    
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
        private void UpdateCOST(string OPCOST, string PCOST, string LINENUM, string DOCENTRY, decimal PRECOST, decimal QTY, decimal PRICE)
        {
            SqlConnection connection = globals.Connection;
            StringBuilder sb = new StringBuilder();
            sb.Append(" UPDATE  SOLAR_PROBOM2 SET OPCOST=@OPCOST,PCOST=@PCOST,PRECOST=@PRECOST,QTY=@QTY,PRICE=@PRICE WHERE SHIPPINGCODE=@SHIPPINGCODE AND DOCENTRY=@DOCENTRY AND LINENUM=@LINENUM");

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            SqlDataAdapter da = new SqlDataAdapter(command);
            command.Parameters.Add(new SqlParameter("@OPCOST", OPCOST));
            command.Parameters.Add(new SqlParameter("@PCOST", PCOST));
            command.Parameters.Add(new SqlParameter("@SHIPPINGCODE", shippingCodeTextBox.Text));
            command.Parameters.Add(new SqlParameter("@LINENUM", LINENUM));
            command.Parameters.Add(new SqlParameter("@DOCENTRY", DOCENTRY));
            command.Parameters.Add(new SqlParameter("@PRECOST", PRECOST));
            command.Parameters.Add(new SqlParameter("@QTY", QTY));
            command.Parameters.Add(new SqlParameter("@PRICE", PRICE));
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
        private void button7_Click(object sender, EventArgs e)
        {

            TOTAL2();
        }
        private void TOTAL2()
        {
            System.Data.DataTable dtCost = MakeTableCombine2();

            System.Data.DataTable DT1 = DT2();
            DataRow dr = null;
            for (int i = 0; i <= DT1.Rows.Count - 1; i++)
            {
                dr = dtCost.NewRow();
                dr["專案代碼"] = DT1.Rows[i]["專案代碼"].ToString().Trim();
                dr["專案名稱"] = DT1.Rows[i]["專案名稱"].ToString().Trim();
                dr["子件編號"] = DT1.Rows[i]["子件編號"].ToString().Trim();
           
                dr["產品名稱"] = DT1.Rows[i]["產品名稱"].ToString().Trim();
                dr["採購成本"] = DT1.Rows[i]["採購成本"].ToString().Trim();
                dr["已付採購成本"] = DT1.Rows[i]["已付採購成本"].ToString().Trim();
                dr["未付採購成本"] = DT1.Rows[i]["未付採購成本"].ToString().Trim();
                dr["期初預估成本"] = DT1.Rows[i]["期初預估成本"].ToString().Trim();
                decimal 採購成本 = Convert.ToDecimal(DT1.Rows[i]["採購成本"].ToString().Trim());
                decimal 已付採購成本 = Convert.ToDecimal(DT1.Rows[i]["已付採購成本"].ToString().Trim());
                decimal 未付採購成本 = Convert.ToDecimal(DT1.Rows[i]["未付採購成本"].ToString().Trim());
                decimal 期初預估成本 = Convert.ToDecimal(DT1.Rows[i]["期初預估成本"].ToString().Trim());

                if (已付採購成本 + 採購成本 == 0)
                {

                    if (期初預估成本 == 0)
                    {
                        dr["預估未付"] = 0;
                    }
                    else
                    {
                        dr["預估未付"] = dr["期初預估成本"];
                    }
                }
                else
                {
                    if (已付採購成本 - 採購成本 == 0)
                    {
                        dr["預估未付"] = "ok";
                    }
                    else
                    {
                        dr["預估未付"] = 採購成本- 已付採購成本;
                    }
                }

                if (採購成本 == 0)
                {

                    dr["實際未來總成本"] = dr["期初預估成本"];

                }
                else
                {
                    dr["實際未來總成本"] = dr["採購成本"];
                }
                //預估未付 實際未來總成本
                dtCost.Rows.Add(dr);
            }

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\SOLAR\\專案成本.xls";

            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            ExcelReport.ExcelReportOutput(dtCost, ExcelTemplate, OutPutFile, "N");

         

        }
        private void button8_Click(object sender, EventArgs e)
        {
            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);
            System.Data.DataTable DT1 = DT21();
            FileName = lsAppDir + "\\Excel\\SOLAR\\專案成本2.xls";

            string ExcelTemplate = FileName;

            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            ExcelReport.ExcelReportOutput(DT1, ExcelTemplate, OutPutFile, "N");

           // ExcelReport.GridViewToExcelDOUBLE(sOLAR_PROBOM2DataGridView, dataGridView1);
        }



    }
}

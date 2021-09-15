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
    public partial class SZ : Form
    {
        string ACME = "";
        string DRS = "";
 
        public SZ()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }

        private System.Data.DataTable GETSZ(string DocEntry, string DocEntry2)
        {

            SqlConnection connection = globals.shipConnection;

            StringBuilder sb = new StringBuilder();
            if (!String.IsNullOrEmpty(DocEntry) && !String.IsNullOrEmpty(DocEntry2))
            {
                sb.Append("              select '進金生' COMPANY,U_RMA_NO RMANO,''''+U_AUO_RMA_NO AUORMANO,U_RVER VER,U_RMODEL MODEL,U_RQUINITY QTY,'' CARGO,U_CUSNAME_S CARDNAME,Contractid,'' P1,'Z19' Z19,'ACME(進金生)' CC,U_RepairCenter REPAIRCENTER,convert(varchar, getdate(), 102) 通知退運日,'1900.01.01' ACME收貨日,'Z0M' Z0M from acmesql02.DBO.OCTR");
                sb.Append("              where Contractid IN (" + DocEntry + ") ");
                sb.Append("              UNION ALL");
                sb.Append("              select '達睿生' COMPANY,U_RMA_NO RMANO,''''+U_AUO_RMA_NO AUORMANO,U_RVER VER,U_RMODEL MODEL,U_RQUINITY QTY,'' CARGO,U_CUSNAME_S CARDNAME,Contractid,'' P1,'Z19' Z19,'ACME(進金生)' CC,U_RepairCenter REPAIRCENTER,convert(varchar, getdate(), 102) 通知退運日,'1900.01.01' ACME收貨日,'Z0M' Z0M from acmesql05.DBO.OCTR");            
                sb.Append("              where Contractid IN (" + DocEntry2 + ") ");
            }
            if (!String.IsNullOrEmpty(DocEntry) && String.IsNullOrEmpty(DocEntry2))
            {
                sb.Append("              select '進金生' COMPANY,U_RMA_NO RMANO,''''+U_AUO_RMA_NO AUORMANO,U_RVER VER,U_RMODEL MODEL,U_RQUINITY QTY,'' CARGO,U_CUSNAME_S CARDNAME,Contractid,'' P1,'Z19' Z19,'ACME(進金生)' CC,U_RepairCenter REPAIRCENTER,convert(varchar, getdate(), 102) 通知退運日,'1900.01.01' ACME收貨日,'Z0M' Z0M from acmesql02.DBO.OCTR");
                sb.Append("              where Contractid IN (" + DocEntry + ") ");
            }
            if (String.IsNullOrEmpty(DocEntry) && !String.IsNullOrEmpty(DocEntry2))
            {
                sb.Append("              select '達睿生' COMPANY,U_RMA_NO RMANO,''''+U_AUO_RMA_NO AUORMANO,U_RVER VER,U_RMODEL MODEL,U_RQUINITY QTY,'' CARGO,U_CUSNAME_S CARDNAME,Contractid,'' P1,'Z19' Z19,'ACME(進金生)' CC,U_RepairCenter REPAIRCENTER,convert(varchar, getdate(), 102) 通知退運日,'1900.01.01' ACME收貨日,'Z0M' Z0M from acmesql05.DBO.OCTR");            
                sb.Append("              where Contractid IN (" + DocEntry2 + ") ");
            }

            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;

       

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "RMA_main");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請選擇");
                    return;
                }
                DELETEFILE();
              

      
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\SZRET.xls";

                TT();
                System.Data.DataTable OrderData = GETSZ(ACME,DRS);
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\AUXM唛头.xls";

                ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");
                string SUBJECT = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_AUXM唛头";
                string DATA = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_AUXM唛头! 请参考～";
                string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                ExcelReport.MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, "", DATA);
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
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp";
                string[] filenames = Directory.GetFiles(OutPutFile);
                foreach (string file in filenames)
                {


                    File.Delete(file);

                }
            }
            catch { }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請選擇");
                    return;
                }
                DELETEFILE();
            
                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\SZRET2.xls";

                TT();
                System.Data.DataTable OrderData = GETSZ(ACME, DRS);

                string ExcelTemplate = FileName;

                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\AUSZ唛头.xls";

                ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");
                string SUBJECT = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_AUSZ唛头";
                string DATA = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_AUSZ唛头! 请参考～";
                string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                ExcelReport.MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, "", DATA);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string strCollected = string.Empty;
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                if (checkedListBox1.GetItemChecked(i))
                {
                    if (strCollected == string.Empty)
                    {
                        strCollected = checkedListBox1.GetItemText(
         checkedListBox1.Items[i]);
                    }
                    else
                    {
                        strCollected = strCollected + checkedListBox1.
         GetItemText(checkedListBox1.Items[i]);
                    }
                }
            }
            string FD = strCollected;
            if (FD == "")
            {
                MessageBox.Show("請選擇公司");
                return;
            }


                 RmaNo frm1 = new RmaNo();
                 frm1.q1 = FD;
                 StringBuilder sb2 = new StringBuilder();
                 StringBuilder sb3 = new StringBuilder();
                 if (frm1.ShowDialog() == DialogResult.OK)
                 {
                     try
                     {
                         string SA = "";
                         string SA2 = "";
                         for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
                         {

                             DataGridViewRow row;

                             row = dataGridView1.Rows[i];
                             string a0 = row.Cells["Contractid"].Value.ToString();
                             string COMPANY = row.Cells["COMPANY"].Value.ToString();
                             if (COMPANY == "進金生")
                             {
                                 sb2.Append(a0 + ",");
                             }
                             if (COMPANY == "達睿生")
                             {
                                 sb3.Append(a0 + ",");
                             }
                         }



                         SA = frm1.q;
                         if (SA.Length > 0)
                         {
                             sb2.Append(SA + ",");
                         }
                         if (sb2.Length > 0)
                         {
                             sb2.Remove(sb2.Length - 1, 1);
                         }

                         SA2 = frm1.q2;
                         if (SA2.Length > 0)
                         {
                             sb3.Append(SA2 + ",");
                         }
                         if (sb3.Length > 0)
                         {
                             sb3.Remove(sb3.Length - 1, 1);
                         }

                         System.Data.DataTable dt1 = GETSZ(sb2.ToString(), sb3.ToString());
                         dataGridView1.DataSource = dt1;
                     }

                     catch (Exception ex)
                     {
                         MessageBox.Show(ex.Message);
                     }
                 }

 

        }

   

        private void SZ_Load(object sender, EventArgs e)
        {
            checkedListBox1.SetItemChecked(0, true);
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請選擇");
                    return;
                }
                DELETEFILE();

                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\不良品退運明細.xls";

                TT();
                System.Data.DataTable OrderData = GETSZ(ACME, DRS);
                //Excel的樣版檔
                string ExcelTemplate = FileName;
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\BV唛头.xls";

                ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");
                string SUBJECT = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_BV唛头";
                string DATA = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_BV唛头! 请参考～";
                string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                ExcelReport.MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, "", DATA);
               
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TT()
        {
            ACME = "";
            DRS = "";
            StringBuilder sb2 = new StringBuilder();
            StringBuilder sb3 = new StringBuilder();
            for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
            {
                DataGridViewRow row;
                row = dataGridView1.SelectedRows[i];
                string a0 = row.Cells["Contractid"].Value.ToString();
                string COMPANY = row.Cells["COMPANY"].Value.ToString();
                if (COMPANY == "進金生")
                {
                    sb2.Append(a0 + ",");
                }
                if (COMPANY == "達睿生")
                {
                    sb3.Append(a0 + ",");
                }
            }
            if (sb2.Length > 0)
            {
                sb2.Remove(sb2.Length - 1, 1);
                ACME = sb2.ToString();
            }
            if (sb3.Length > 0)
            {
                sb3.Remove(sb3.Length - 1, 1);
                DRS = sb3.ToString();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {

            for (int i2 = 0; i2 <= dataGridView1.Rows.Count - 1; i2++)
            {
                 
                DataGridViewRow row;

                row = dataGridView1.Rows[i2];
                //COMPANY
                string COMPANY = row.Cells["COMPANY"].Value.ToString();
                string RMANO = row.Cells["RMANO"].Value.ToString();
                string REPAIRCENTER = row.Cells["REPAIRCENTER"].Value.ToString();
                DateTime 通知退運日 = Convert.ToDateTime(row.Cells["通知退運日"].Value);
                DateTime ACME收貨日 = Convert.ToDateTime(row.Cells["ACME收貨日"].Value);

                if (COMPANY == "進金生")
                {
                    UPDATEJOBNO(REPAIRCENTER, 通知退運日, ACME收貨日, RMANO);
                }

                if (COMPANY == "達睿生")
                {
                    UPDATEJOBNO5(REPAIRCENTER, 通知退運日, ACME收貨日, RMANO);
                }
            }
        }

        public void UPDATEJOBNO(string U_RepairCenter, DateTime U_Racmetodate, DateTime U_RtoReceiving, string U_RMA_NO)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE acmesql02.DBO.OCTR SET U_RepairCenter=@U_RepairCenter,U_Racmetodate=@U_Racmetodate,U_RtoReceiving=@U_RtoReceiving WHERE U_RMA_NO =@U_RMA_NO", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RepairCenter", U_RepairCenter));
            command.Parameters.Add(new SqlParameter("@U_Racmetodate", U_Racmetodate));
            command.Parameters.Add(new SqlParameter("@U_RtoReceiving", U_RtoReceiving));
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
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

        public void UPDATEJOBNO5(string U_RepairCenter, DateTime U_Racmetodate, DateTime U_RtoReceiving, string U_RMA_NO)
        {
            SqlConnection connection = new SqlConnection(globals.shipConnectionString);
            SqlCommand command = new SqlCommand("UPDATE acmesql05.DBO.OCTR SET U_RepairCenter=@U_RepairCenter,U_Racmetodate=@U_Racmetodate,U_RtoReceiving=@U_RtoReceiving WHERE U_RMA_NO =@U_RMA_NO", connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RepairCenter", U_RepairCenter));
            command.Parameters.Add(new SqlParameter("@U_Racmetodate", U_Racmetodate));
            command.Parameters.Add(new SqlParameter("@U_RtoReceiving", U_RtoReceiving));
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", U_RMA_NO));
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
            try
            {
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請選擇");
                    return;
                }
                DELETEFILE();



                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\SZRET普倉.xls";

                TT();
                System.Data.DataTable OrderData = GETSZ(ACME, DRS);
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\DRS普仓唛头.xls";

                ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");
                string SUBJECT = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_DRS普仓唛头";
                string DATA = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_DRS普仓唛头! 请参考～";
                string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                ExcelReport.MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, "", DATA);
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
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請選擇");
                    return;
                }
                DELETEFILE();



                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\SZRETZ0M.xls";

                TT();
                System.Data.DataTable OrderData = GETSZ(ACME, DRS);
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\Z0M唛头.xls";

                ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");
                string SUBJECT = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_Z0M唛头";
                string DATA = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_Z0M唛头! 请参考～";
                string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                ExcelReport.MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, "", DATA);
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
                if (dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("請選擇");
                    return;
                }
                DELETEFILE();



                string FileName = string.Empty;
                string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                FileName = lsAppDir + "\\Excel\\RMA\\SZRETHSO.xls";

                TT();
                System.Data.DataTable OrderData = GETSZ(ACME, DRS);
                //Excel的樣版檔
                string ExcelTemplate = FileName;

                //輸出檔
                string OutPutFile = lsAppDir + "\\Excel\\RMA\\temp\\HSO唛头.xls";

                ExcelReport.ExcelReportOutputLEMON(OrderData, ExcelTemplate, OutPutFile, "N");
                string SUBJECT = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_HSO唛头";
                string DATA = DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("dd") + "_HSO唛头! 请参考～";
                string MAIL = fmLogin.LoginID.ToString() + "@acmepoint.com";
                ExcelReport.MailTest(SUBJECT, fmLogin.LoginID.ToString(), MAIL, "", DATA);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

      
    }
}
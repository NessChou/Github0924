using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;



namespace ACME
{
    public partial class RMAODLN : Form
    {
        public RMAODLN()
        {
            InitializeComponent();
        }


        private void RMAODLN_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'rm.PARAMS' 資料表。您可以視需要進行移動或移除。
            this.pARAMSTableAdapter.Fill(this.rm.PARAMS);

            System.Data.DataTable dt4 = GetMenu.GETRMAWH();


            comboBox1.Items.Clear();


            for (int i = 0; i <= dt4.Rows.Count - 1; i++)
            {
                comboBox1.Items.Add(Convert.ToString(dt4.Rows[i][0]));
            }
        }

        private void pARAMSDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["PARAM_KIND"].Value = "RMAWH";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.pARAMSBindingSource.EndEdit();
            this.pARAMSTableAdapter.Update(this.rm.PARAMS);

        }

        private void button2_Click(object sender, EventArgs e)
        {



                string NumberName = "RMN" + DateTime.Now.ToString("yyyyMMdd");
                string NumberName2 = "RMA" + DateTime.Now.ToString("yyyyMMdd");
                string AutoNum = util.GetAutoNumber(globals.Connection, NumberName);
                string F1 = comboBox1.Text + "收貨通知單---" + NumberName2 + AutoNum + "X";


                try
                {
                    string FileName = string.Empty;
                    string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

                    FileName = lsAppDir + "\\Excel\\RMA\\收貨工單.xls";


                    //Excel的樣版檔
                    string ExcelTemplate = FileName;

                    //輸出檔
                    string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                          DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

                    //產生 Excel Report
                    ExcelReport.ExcelReportOutput(GetOrderData(F1), ExcelTemplate, OutPutFile, "N");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
        }

        private System.Data.DataTable GetOrderData(string F1)
        {

            SqlConnection connection = globals.shipConnection ;

            StringBuilder sb = new StringBuilder();

            sb.Append(" SELECT ''''+U_AUO_RMA_NO VRMANO,ROW_NUMBER() OVER(ORDER BY U_RMA_NO) AS  LINE,'PCS' PCS,CONVERT(varchar(100), GETDATE(), 111) DNOW,U_RVender VENDER,U_RMA_NO RMANO,U_Cusname_S CUST");
            sb.Append(" ,U_RMODEL MODEL,U_RVer VER,U_RGRADE GRADE,U_Rquinity QTY,'' INVOICE,");
            sb.Append(" (SELECT TOP 1 U_U_PLACE_1    FROM CTR1 WHERE ContractID =T0.ContractID AND ISNULL(U_U_PLACE_1,'') <> '') LOCATION,");
            sb.Append(" @TITLE TITLE,@LGOIN LGOIN  FROM OCTR  T0 WHERE U_RMA_NO in ( " + textBox1.Text + ") ORDER BY   U_RMA_NO ");
            SqlCommand command = new SqlCommand(sb.ToString(), connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@U_RMA_NO", textBox1.Text));
            command.Parameters.Add(new SqlParameter("@TITLE", F1));
            command.Parameters.Add(new SqlParameter("@LGOIN", fmLogin.LoginID.ToString()));

            SqlDataAdapter da = new SqlDataAdapter(command);

            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "rma_PackingListM");
            }
            finally
            {
                connection.Close();
            }

            return ds.Tables[0];

        }
    }
}

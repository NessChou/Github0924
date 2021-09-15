using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ACME
{

    public partial class ADOQUT : Form
    {
        public string q;


        public ADOQUT()
        {
            InitializeComponent();
        }

        private void ADOQUT_Load(object sender, EventArgs e)
        {
            // TODO: 這行程式碼會將資料載入 'aD_OQUT._AD_OQUT' 資料表。您可以視需要進行移動或移除。
            this.aD_OQUTTableAdapter.Fill(this.aD_OQUT._AD_OQUT);
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {

        }

        private void ADOQUTBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //this.aD_OQUTTableAdapter.Fill(this.aD_OQUT._AD_OQUT);
            this.aD_OQUTTableAdapter.FillBy(this.aD_OQUT._AD_OQUT, textIRFNO.Text, textVENDOR.Text, textSIZE.Text);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textIRFNO.Text = "";
            textVENDOR.Text = "";
            textSIZE.Text = "";
            textCUSNAME.Text = "";
            this.aD_OQUTTableAdapter.FillBy(this.aD_OQUT._AD_OQUT, textIRFNO.Text, textVENDOR.Text, textSIZE.Text);
        }


        private void _1218testTableBindingNavigatorSaveItem_Click_1(object sender, EventArgs e)
        {

        }

        private void _1218testTableBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {

        }

        private void BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.aDOQUTBindingSource.EndEdit();
            this.aD_OQUTTableAdapter.Update(this.aD_OQUT);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (aD_OQUTDataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("請點選單號");
                return;
            }

            ArrayList al = new ArrayList();
            for (int i = 0; i <= listBox1.Items.Count - 1; i++)
            {
                al.Add(listBox1.Items[i].ToString());
            }
            StringBuilder sb = new StringBuilder();
            foreach (string v in al)
            {
                sb.Append("'" + v + "',");
            }
            q = sb.Remove(sb.Length - 1, 1).ToString();
            

            string FileName = string.Empty;
            string lsAppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName);

            FileName = lsAppDir + "\\Excel\\OPCH\\採購報價.xls";

            //輸出檔
            string OutPutFile = lsAppDir + "\\Excel\\temp\\" +
                  DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetFileName(FileName);

            System.Data.DataTable G1 = null;

            G1 = GetEXCEL(q);

            ExcelReport.ExcelReportOutput(G1, FileName, OutPutFile, "N");
        }

        private void aD_OQUTDataGridView_MouseCaptureChanged(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                if (aD_OQUTDataGridView.SelectedRows.Count > 0)
                {
                    DataGridViewRow row;
                    for (int i = aD_OQUTDataGridView.SelectedRows.Count - 1; i >= 0; i--)
                    {
                        row = aD_OQUTDataGridView.SelectedRows[i];
                        string row123 = aD_OQUTDataGridView.SelectedRows[i].Cells[0].Value.ToString();
                        listBox1.Items.Add(row123);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public System.Data.DataTable GetEXCEL(string cs)
        {
            SqlConnection MyConnection = globals.Connection;

            StringBuilder sb = new StringBuilder();
            sb.Append("  SELECT ID, RFQ_WORKSHEET, ");
            sb.Append("  		CTRL_CODE, SIZE, ");
            sb.Append("  		MODEL, VENDOR, Reference, ");
            sb.Append("  		LCD_PANEL, DISPLAY_MODE, ");
            sb.Append("  		RESOLUTION, BRIGHTNESS, ");
            sb.Append("  		VIEWING_ANGLE, TOUCH_PANNEL, ");
            sb.Append("  		TEMPERED_GLASS, IO_PORT, ");
            sb.Append("  		AC_POWER_BOARD, POWER_TYPE, ");
            sb.Append("  		DC_CONVERTER, TOUCH_CONTROLLER, ");
            sb.Append("  		BACKLIGHT_DRIVER, OSD_CONTROL, ");
            sb.Append("  		SPEAKER, FAN, POWER_CORD, ");
            sb.Append("  		ACCESSORY, COLOR, FRONT_COVER, ");
            sb.Append("  		BACK_COVER, FORM_FACTOR, ");
            sb.Append("  		MOUNTING, ORIENTATION, ");
            sb.Append("  		NOTE, TOUCH, NRE, PANEL_SAMPLE, ");
            sb.Append("  		PANEL_50PCS, PANEL_500PCS, ");
            sb.Append("  		PANEL_1000PCS, KIT_SAMPLE, ");
            sb.Append("  		KIT_50PCS, KIT_500PCS, ");
            sb.Append("  		KIT_1000PCS, REMARKS ");
            sb.Append("  FROM dbo.[AD_OQUT] ");
            sb.Append("  WHERE ID in ( " + cs + ") ");

            SqlCommand command = new SqlCommand(sb.ToString(), MyConnection);
            command.CommandType = CommandType.Text;

            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();

            try
            {
                MyConnection.Open();
                da.Fill(ds, "APLC");
            }
            finally
            {
                MyConnection.Close();
            }
            return ds.Tables["APLC"];
        }
    }
}

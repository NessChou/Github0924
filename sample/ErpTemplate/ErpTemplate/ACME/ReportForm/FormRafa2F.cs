using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.Shared;
using System.Drawing.Printing;
namespace ACME
{

    public partial class FormRafa2F : Form
    {
        public DataTable dt;
        public FormRafa2F()
        {
            InitializeComponent();
        }

        private void Form1Rpt5_Load(object sender, EventArgs e)
        {
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;


            logOnInfo.ConnectionInfo.ServerName = "acmesap";
            logOnInfo.ConnectionInfo.DatabaseName = "acmesqlSP";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
            Form1Cry71.Database.Tables[i].ApplyLogOnInfo(logOnInfo);

            Form1Cry71.SetDataSource(dt);
            crystalReportViewer1.ReportSource = Form1Cry71;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintDialog MyPDI = new PrintDialog(); //列印對話框
            PrintDocument MyPDO = new PrintDocument();//列印文件

            MyPDI.Document = MyPDO;

            if (MyPDI.ShowDialog() == DialogResult.OK)
            {
                Form1Cry71.PrintOptions.PrinterName = MyPDI.PrinterSettings.PrinterName;
                Form1Cry71.PrintToPrinter(MyPDO.PrinterSettings.Copies, true, 0, 0);
            }
        }
    }
}
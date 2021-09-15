using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.Shared;

namespace ACME
{
    public partial class Form1Rpt : Form
    {
        public DataTable dt;
        public Form1Rpt()
        {
            InitializeComponent();
        }

        private void Form1Rpt_Load(object sender, EventArgs e)
        {
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;

            logOnInfo.ConnectionInfo.ServerName = "acmesap";
            logOnInfo.ConnectionInfo.DatabaseName = "acmesql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
            CrystalReport11.Database.Tables[i].ApplyLogOnInfo(logOnInfo);
            CrystalReport11.SetDataSource(dt);
            crystalReportViewer1.ReportSource = CrystalReport11;
        }
    }
}
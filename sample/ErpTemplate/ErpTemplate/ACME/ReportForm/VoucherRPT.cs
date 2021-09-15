using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System.Data.SqlClient;
namespace ACME
{
    public partial class VoucherRPT : Form
    {
        public DataTable dt;
        public VoucherRPT()
        {
            InitializeComponent();
        }

       

        private void OvpmRpt_Load(object sender, EventArgs e)
        {
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;

            logOnInfo.ConnectionInfo.ServerName = "acmesap";
            logOnInfo.ConnectionInfo.DatabaseName = "acmesql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
            OVPMlReport11.Database.Tables[i].ApplyLogOnInfo(logOnInfo);
            OVPMlReport11.SetDataSource(dt);
            crystalReportViewer1.ReportSource = OVPMlReport11;
        }
    }
}
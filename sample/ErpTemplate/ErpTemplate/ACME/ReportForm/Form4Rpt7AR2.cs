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
    public partial class Form4Rpt7AR2 : Form
    {
        public Form4Rpt7AR2()
        {
            InitializeComponent();
        }
        public string s;
        public DataTable dt;

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
    
        }

        private void Form4Rpt7_Load(object sender, EventArgs e)
        {

 
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;

            logOnInfo.ConnectionInfo.ServerName = "acmesap";
            logOnInfo.ConnectionInfo.DatabaseName = "acmesql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
           

            APSCry33.Database.Tables[i].ApplyLogOnInfo(logOnInfo);
            APSCry33.SetDataSource(dt);
            crystalReportViewer2.ReportSource = APSCry33;
        }


    }
}
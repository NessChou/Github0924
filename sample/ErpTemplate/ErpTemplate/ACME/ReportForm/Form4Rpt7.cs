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
    public partial class Form4Rpt7 : Form
    {
        public Form4Rpt7()
        {
            InitializeComponent();
        }
        public string s;
        public DataTable dt;
        public DataTable dt2;
        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
    
        }

        private void Form4Rpt7_Load(object sender, EventArgs e)
        {

 
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
      

            logOnInfo.ConnectionInfo.ServerName = "acmesap";
            logOnInfo.ConnectionInfo.DatabaseName = "acmesql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";


            APSCry33.Database.Tables[0].ApplyLogOnInfo(logOnInfo);
            APSCry33.Database.Tables[0].SetDataSource(dt);
            APSCry33.Subreports[0].Database.Tables[0].ApplyLogOnInfo(logOnInfo);
            APSCry33.Subreports[0].Database.Tables[0].SetDataSource(dt2);
            crystalReportViewer2.ReportSource = APSCry33;
        }


    }
}
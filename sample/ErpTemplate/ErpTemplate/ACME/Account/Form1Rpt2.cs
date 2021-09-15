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
    public partial class Form1Rpt2 : Form
    {
        public DataTable dt;
        public Form1Rpt2()
        {
            InitializeComponent();
        }

        private void Form1Rpt2_Load(object sender, EventArgs e)
        {
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;

            // 在報表的所有資料表中進行迴圈。
            //for (i = 0; i == CrystalReport11.Database.Tables.Count - 1; i++)
            //{
            // 設定目前資料表的連接資訊。
            logOnInfo.ConnectionInfo.ServerName = "acmesrv13";
            logOnInfo.ConnectionInfo.DatabaseName = "acmesql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
            CrystalReport11.Database.Tables[i].ApplyLogOnInfo(logOnInfo);
            //    }
            // = false;
            CrystalReport11.SetDataSource(dt);
            crystalReportViewer1.ReportSource = CrystalReport11;
        }
    }
}
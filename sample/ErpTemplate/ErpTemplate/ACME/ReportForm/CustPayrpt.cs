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
    public partial class CustPayrpt : Form
    {
        public string s, q, money;
        public CustPayrpt()
        {
            InitializeComponent();
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
            ParameterFields pfs = new ParameterFields();

            //宣告個別參數
            ParameterField pf1 = new ParameterField();

            //設定參數名稱
            pf1.Name = "YM";

 

            //宣告參數值
            ParameterDiscreteValue pdv1 = new ParameterDiscreteValue();

            //設定參數值
            pdv1.Value = q;
            pf1.CurrentValues.Add(pdv1);
            //加入參數集合
            pfs.Add(pf1);




            crystalReportViewer1.ParameterFieldInfo = pfs;
        }

        private void CustPayrpt_Load(object sender, EventArgs e)
        {
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;


            logOnInfo.ConnectionInfo.ServerName = "acmesap";
            logOnInfo.ConnectionInfo.DatabaseName = "AcmeSql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
            CustReport11.Database.Tables[i].ApplyLogOnInfo(logOnInfo);

            crystalReportViewer1.ReportSource = CustReport11;
        }
    }
}
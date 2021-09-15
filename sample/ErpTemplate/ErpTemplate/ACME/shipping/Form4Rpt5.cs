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
    public partial class Form4Rpt5 : Form
    {
        public Form4Rpt5()
        {
            InitializeComponent();
        }
        public string s, q, es, w;  

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
            ParameterFields pfs = new ParameterFields();

            //宣告個別參數
            ParameterField pf1 = new ParameterField();

            ParameterField pf4 = new ParameterField();
            ParameterField pf5 = new ParameterField();
            //設定參數名稱
            pf1.Name = "YM";

            pf4.Name = "YM4";
            pf5.Name = "YM5";
            //宣告參數值
            ParameterDiscreteValue pdv1 = new ParameterDiscreteValue();


            ParameterDiscreteValue pdv4 = new ParameterDiscreteValue();
            ParameterDiscreteValue pdv5 = new ParameterDiscreteValue();
            //設定參數值
            pdv1.Value = s;
            pf1.CurrentValues.Add(pdv1);
            //加入參數集合
            pfs.Add(pf1);


            pdv4.Value = w;
            pf4.CurrentValues.Add(pdv4);
            //加入參數集合
            pfs.Add(pf4);

            pdv5.Value = q;
            pf5.CurrentValues.Add(pdv5);
            //加入參數集合
            pfs.Add(pf5);
            // ConfigureCrystalReports();
            //報表設定參數集合
            crystalReportViewer1.ParameterFieldInfo = pfs;
        }

        private void Form4Rpt5_Load(object sender, EventArgs e)
        {
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;


            logOnInfo.ConnectionInfo.ServerName = "acmesrv13";
            logOnInfo.ConnectionInfo.DatabaseName = "AcmeSql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
            CrystalReport31.Database.Tables[i].ApplyLogOnInfo(logOnInfo);

            crystalReportViewer1.ReportSource = CrystalReport31;
        }


    }
}
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
    public partial class ShipInsuRpt : Form
    {
        public string s, q;
        public ShipInsuRpt()
        {
            InitializeComponent();
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
            ParameterFields pfs = new ParameterFields();

            //宣告個別參數
            ParameterField pf1 = new ParameterField();
       
            ParameterField pf4 = new ParameterField();

            //設定參數名稱
            pf1.Name = "YM";
 
            pf4.Name = "YM4";
           // pf5.Name = "YM5";
            //宣告參數值
            ParameterDiscreteValue pdv1 = new ParameterDiscreteValue();

    
           ParameterDiscreteValue pdv4 = new ParameterDiscreteValue();

            //設定參數值
            pdv1.Value = s;
            pf1.CurrentValues.Add(pdv1);
            //加入參數集合
            pfs.Add(pf1);


            pdv4.Value = q;
            pf4.CurrentValues.Add(pdv4);
            //加入參數集合
            pfs.Add(pf4);


           // ConfigureCrystalReports();
            //報表設定參數集合
            crystalReportViewer1.ParameterFieldInfo = pfs;

        }

        private void Form2Rpt_Load(object sender, EventArgs e)
        {
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;


            logOnInfo.ConnectionInfo.ServerName = "acmesrv13";
            logOnInfo.ConnectionInfo.DatabaseName = "AcmeSql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
            Form5Cry1.Database.Tables[i].ApplyLogOnInfo(logOnInfo);

            crystalReportViewer1.ReportSource = Form5Cry1;
        }

    }
}
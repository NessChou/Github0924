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
    public partial class Form1RptACC : Form
    {
        public string s,q;
        public DataTable dt;
        public Form1RptACC()
        {
            InitializeComponent();
        }

        private void Form1Rpt4_Load(object sender, EventArgs e)
        {

            ParameterFields pfs = new ParameterFields();

            //宣告個別參數
            ParameterField pf1 = new ParameterField();
            ParameterField pf2 = new ParameterField();
            //設定參數名稱
            pf1.Name = "YM";
            pf2.Name = "YM1";
            //宣告參數值
            ParameterDiscreteValue pdv1 = new ParameterDiscreteValue();
            ParameterDiscreteValue pdv2 = new ParameterDiscreteValue();

            //設定參數值
            pdv1.Value = s;
            pf1.CurrentValues.Add(pdv1);
            //加入參數集合
            pfs.Add(pf1);

            pdv2.Value = q;
            pf2.CurrentValues.Add(pdv2);
            //加入參數集合
            pfs.Add(pf2);

            //報表設定參數集合
            crystalReportViewer1.ParameterFieldInfo = pfs;


            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;

            // 在報表的所有資料表中進行迴圈。
            //for (i = 0; i == CrystalReport11.Database.Tables.Count - 1; i++)
            //{
            // 設定目前資料表的連接資訊。
            logOnInfo.ConnectionInfo.ServerName = "acmesap";
            logOnInfo.ConnectionInfo.DatabaseName = "acmesql02";
            logOnInfo.ConnectionInfo.UserID = "sapdbo";
            logOnInfo.ConnectionInfo.Password = "@rmas";
            Form1Cry41.Database.Tables[i].ApplyLogOnInfo(logOnInfo);
            //    }
            // = false;
            Form1Cry41.SetDataSource(dt);
            crystalReportViewer1.ReportSource = Form1Cry41;
        }
    }
}
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
    public partial class fmRmarptshow : Form
    {
        public string aa;
        public fmRmarptshow()
        {
            InitializeComponent();
        }

        private void crystalReportViewer1_Load(object sender, EventArgs e)
        {
                    
        

            //宣告參數集合
            ParameterFields pfs = new ParameterFields();

            //宣告個別參數
            ParameterField pf1 = new ParameterField();
         

            //設定參數名稱
            pf1.Name = "HCUST_KEY";
         

            //宣告參數值
            ParameterDiscreteValue pdv1 = new ParameterDiscreteValue();
          
            //設定參數值
            pdv1.Value = aa;
            pf1.CurrentValues.Add(pdv1);
            //加入參數集合
            pfs.Add(pf1);

          

            //   ConfigureCrystalReports();
            //報表設定參數集合
            crystalReportViewer1.ParameterFieldInfo = pfs;
           
        }

        

    }
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME.CRM
{
    public partial class fmRptStage : Form
    {
        public DataTable dt;
        
        public fmRptStage()
        {
            InitializeComponent();
        }

        private void fmRptStage_Load(object sender, EventArgs e)
        {
            crmStage1.SetDataSource(dt);

            crystalReportViewer1.ReportSource = crmStage1;
        }
    }
}
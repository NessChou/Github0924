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
    public partial class FrmPrint : Form
    {
        public FrmPrint()
        {
            InitializeComponent();
        }

        public DataSet ds;
        public DataTable dt;

        private void FrmPrint_Load(object sender, EventArgs e)
        {
            TableLogOnInfo logOnInfo = new TableLogOnInfo();
            int i = 0;

            // �b�����Ҧ���ƪ��i��j��C
            //for (i = 0; i == CrystalReport11.Database.Tables.Count - 1; i++)
            //{
                // �]�w�ثe��ƪ��s����T�C
            logOnInfo.ConnectionInfo.ServerName = "acmesrv13";
            logOnInfo.ConnectionInfo.DatabaseName = "AcmeSql02";
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
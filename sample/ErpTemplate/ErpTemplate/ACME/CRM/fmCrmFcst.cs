using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;

namespace ACME.CRM
{
    public partial class fmCrmFcst : Form
    {
        public DataTable dt;
        
        public fmCrmFcst()
        {
            InitializeComponent();
        }

        private void fmCrmFcst_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = dt;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(Application.StartupPath);
             
           // WebClient myWebClient=new WebClient();
        
            //Application.DoEvents();

            //string RemoteUri =@"http://portal.acmepoint.net/rma/AcmeFiles/WebCRM/" ;
            //string DownloadExeFile = "AutoUpdateWeb.exe";
            //string MainExe = "Acme.exe";
            //string ManifestFile = "ServerManifest.xml";

            //string UserID;

            //string Web = "WEB";

            //string AcmeDirectory = @"C:\AcmeCRM\";

            //string strOK = "Y";

            //if (!Directory.Exists(AcmeDirectory))
            //{
            //    DirectoryInfo di = Directory.CreateDirectory(AcmeDirectory);

            //}

            //try
            //{
            //    this.Cursor = Cursors.WaitCursor;

            //    // AutoUpdate
            //    if (!(System.IO.File.Exists(AcmeDirectory + DownloadExeFile)))
            //    {

            //        myWebClient.DownloadFile(RemoteUri + DownloadExeFile, AcmeDirectory + DownloadExeFile);
            //    }
        }

      
    }
}
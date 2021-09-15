using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME.CRM
{
    public partial class CrmMis : Form
    {
        public CrmMis()
        {
            InitializeComponent();
        }

        public string UserId;

        private void CrmMis_Load(object sender, EventArgs e)
        {
            string sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='IssueKind'";
            UtilSimple.SetLookupBinding(comboBox1, globals.Connection, sql, aCME_MISBindingSource, "IssueKind",
            "PARAM_DESC", "PARAM_DESC");

            sql = "SELECT *  FROM ACME_PARAMS WHERE PARAM_KIND='ActionFlag'";
            UtilSimple.SetLookupBinding(comboBox2, globals.Connection, sql, aCME_MISBindingSource, "ActionFlag",
            "PARAM_DESC", "PARAM_DESC");


        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void exBindingNavigator1_AfterNew(object sender, EventArgs e)
        {
            userCodeTextBox.Text = UserId;
            docDateTextBox.Text = DateTime.Now.ToString("yyyyMMdd");
            //給值
           // ((DataRowView)aCME_MISBindingSource.Current).Row["UserCode"] = UserId;
            ((DataRowView)aCME_MISBindingSource.Current).Row["DocNum"] = DateTime.Now.ToString("yyyyMMddhhmmssss");
   
        }

        private void exBindingNavigator1_BeforeDelete(object sender, MyEventArgs args)
        {

            args.CheckOk = false;

            if (UserId.ToUpper() == "TERRYLEE" ||
                UserId.ToUpper() == "LLEYTONCHEN" ||
                UserId.ToUpper() == "ANNIECHEN")
            {

                args.CheckOk = true;

            }
        }
    }
}
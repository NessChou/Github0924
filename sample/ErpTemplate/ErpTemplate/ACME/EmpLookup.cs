using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using YYJXC.Service;

namespace YYJXC.BaseForm
{
    public partial class EmpLookup : YYJXC.LookupDialog
    {
        public EmpLookup()
        {
            InitializeComponent();
        }

        private DataSet dataSet1;

        private string _selectId;

        public string SelectID
        {
            get
            {
                return _selectId;
            }
        }

        private string _selectName;

        public string SelectName
        {
            get
            {
                return _selectName;
            }
        }

        protected override void ProcessOK()
        {
            if (BindDataSource.Current != null)
            {
                _selectId = (string)((DataRowView)((BindingSource)BindDataSource).Current)["EMP_NO"];
                _selectName = (string)((DataRowView)((BindingSource)BindDataSource).Current)["EMP_NAME"];
            }
        }

        protected override void InitFields(ComboBox cb)
        {
            base.InitFields(cb);

            cb.Items.Add("員工編號(EMP_NO)");
            cb.Items.Add("姓名(EMP_NAME)");
            cb.SelectedIndex = 0;
           
        }

      
        private void EmpLookup_Load(object sender, EventArgs e)
        {
            

            ViewData();

            InitializeForm(bindingSource1);
        }

        public void ViewData()
        {
            
            this.Cursor = Cursors.WaitCursor;

            // Create the proxy.
            WebService1 proxy = new WebService1();
            proxy.Url = frmLoad.HostName;

           dataSet1 = proxy.GetEMPData(frmLoad.LoginID, frmLoad.PWD, "1");

     
           
            this.Cursor = Cursors.Default;
        }


        protected override void InitializeForm(BindingSource dataSource)
        {
            base.InitializeForm(dataSource);
            dataSource.DataSource = dataSet1.Tables[0];
            dataGridView1.DataSource = dataSource;
            dataGridView1.AutoGenerateColumns = false;
            LookupDataset = dataSet1;
        }

      

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            btnOK_Click(sender, EventArgs.Empty);
        }

    }
}


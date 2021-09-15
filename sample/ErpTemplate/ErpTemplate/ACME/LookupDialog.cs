using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace YYJXC
{
	public partial class LookupDialog : Form
	{
		private BindingSource _dataSource;
        public DataSet LookupDataset;


		public BindingSource BindDataSource
		{
			get
			{
				return _dataSource;
			}
            set
            {
                _dataSource = value;
            }
		}

		public LookupDialog()
		{
			InitializeComponent();
		}

        protected virtual void InitializeForm(BindingSource dataSource)
		{
            //gridLookup.AutoGenerateColumns = false;
            //gridLookup.DataSource = dataSource;
            //_dataSource = dataSource;
          //  BindDataSource = new BindingSource();
            BindDataSource = dataSource;
            
		}

        //執行
		protected virtual void ProcessOK()
		{
		}

        //增加查詢欄位
        protected virtual void InitFields(ComboBox cb)
        {
            cbFields = cb;
        }
        
        //
        protected  void DoQuery(DataTable dt)
        {
            DataView dv = new DataView(dt);

            if (tbExpression.Text.Trim() == "")
            {
                MessageBox.Show("請輸入查詢條件值！", "信息提示");
                tbExpression.Focus();
                return;
            }
            try
            {

                if (cbOperator.Text == "like")
                {

                    dv.RowFilter =GetFieldName(cbFields.Text) + " like '*" + tbExpression.Text.Trim() + "*'";
                    BindDataSource.DataSource = dv;

                }
                else
                {

                    dv.RowFilter = GetFieldName(cbFields.Text) + cbOperator.Text + "'" + tbExpression.Text.Trim() + "'";
                    BindDataSource.DataSource = dv;
                }

            }
            catch (Exception err)
            {

                MessageBox.Show("操作出現錯誤：" + err.Message, "信息提示", MessageBoxButtons.OK, MessageBoxIcon.Error);
                cbFields.Focus();
            }

        }

        private string GetFieldName(string s)
        {
            int i =s.IndexOf("(");
            string t =s.Substring(i+1, s.Length - i-2);
            return t;
        }

		private void btnCancel_Click(object sender, EventArgs e)
		{
			Close();
		}

		public void btnOK_Click(object sender, EventArgs e)
		{
			this.DialogResult = DialogResult.OK;
            ProcessOK();
			Close();
		}

		private void gridLookup_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
		{
			btnOK_Click(this, EventArgs.Empty);
		}

        private void LookupDialog_Load(object sender, EventArgs e)
        {
            cbOperator.Items.Add("like");
            cbOperator.Items.Add("=");
            cbOperator.Items.Add(">");
            cbOperator.Items.Add("<");
            cbOperator.SelectedIndex = 0;
            InitFields(cbFields);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DoQuery(LookupDataset.Tables[0]);
        }
       
	}
}
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ACME
{
	public partial class UserValueDialog : Form
	{
		private BindingSource _dataSource;
        public DataTable LookupDataset;


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


        private string fFormID;

        public string FormID1
        {
            get
            {
                return fFormID;
            }
            set
            {
                fFormID = value;
            }

        }

        private string fObjID;

        public string ObjID1
        {
            get
            {
                return fObjID;
            }
            set
            {
                fObjID = value;
            }

        }


        private string _selectValue;

        public string SelectValue
        {
            get
            {
                return _selectValue;
            }
            set
            {
                _selectValue = value;
            }
        }

		public UserValueDialog()
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
            DataGridViewRow drv = aCME_UserValueDataGridView.CurrentRow;

            SelectValue = Convert.ToString(drv.Cells[2].Value);
          

		}

        //增加查詢欄位
        protected virtual void InitFields(ComboBox cb)
        {
            
        }
        
        //
        protected  void DoQuery(DataTable dt,string FieldName,string Value)
        {
            //for (int i = 0; i <= dt.Rows.Count - 1; i++)
            //{ 
               
            
            //}

            

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
			
		}

        private void LookupDialog_Load(object sender, EventArgs e)
        {
            try
            {
                //加解密後的Connection
                aCME_UserValueTableAdapter.Connection = globals.Connection;
                aCME_UserValueTableAdapter.Fill(userValue.ACME_UserValue,FormID1,ObjID1);


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

       

        private void aCME_UserValueDataGridView_DefaultValuesNeeded(object sender, DataGridViewRowEventArgs e)
        {
            e.Row.Cells["FormID"].Value = FormID1;
            e.Row.Cells["ObjID"].Value = ObjID1;

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            
            try
            {
                aCME_UserValueTableAdapter.Update(userValue.ACME_UserValue);
                MessageBox.Show("存檔成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void aCME_UserValueDataGridView_DoubleClick(object sender, EventArgs e)
        {
            btnOK_Click(this, EventArgs.Empty);
        }

        private void tbExpression_TextChanged(object sender, EventArgs e)
        {
            //尋找
            //Key 值要完全相同
            //aCME_UserValueBindingSource.Position = aCME_UserValueBindingSource.Find("KeyValue", tbExpression.Text);
            if (tbExpression.Text.Length > 0)
            {
                int loc = aCME_UserValueBindingSource.Find("KeyValue", tbExpression.Text);
                if (loc == -1)
                {
                    DataView dv = new DataView(userValue.ACME_UserValue);
                    dv.RowFilter = string.Format("KeyValue LIKE '{0}*'",tbExpression.Text);
                    if (dv.Count > 0)
                    {
                        loc = aCME_UserValueBindingSource.Find("KeyValue", dv[0]["KeyValue"]);
                    }
                }
                aCME_UserValueBindingSource.Position = loc;
            }

        }
       
	}
}
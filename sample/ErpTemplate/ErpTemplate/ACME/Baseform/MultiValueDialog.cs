using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace ACME
{
	public partial class MultiValueDialog : Form
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

        private DataTable dataSet1;

        private DataTable fSourceDataTable;


        public DataTable SourceDataTable
        {
            get
            {
                return fSourceDataTable;
            }
            set
            {
                fSourceDataTable = value;
            }
        }


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

        private string[] fFieldNames;

        public string[] FieldNames
        {
            get
            {
                return fFieldNames;
            }
            set
            {
                fFieldNames = value;
            }

        }


        private string[] fCaptions;

        public string[] Captions
        {
            get
            {
                return fCaptions;
            }
            set
            {
                fCaptions = value;
            }

        }

        private string fSqlScript;

        public string SqlScript
        {
            get
            {
                return fSqlScript;
            }
            set
            {
                fSqlScript = value;
            }
        }


        private object[] fLookupValues;

        public object[] LookupValues
        {
            get
            {
                return fLookupValues;
            }
            set
            {
                fLookupValues = value;
            }
        }

        private SqlConnection fConnection;

        public SqlConnection LookUpConnection
        {
            get
            {
                return fConnection;
            }
            set
            {
                fConnection = value;
            }
        }

        //???w?^??????
        private string fKeyFieldName;

        public string KeyFieldName
        {
            get
            {
                return fKeyFieldName;
            }
            set
            {
                fKeyFieldName = value;
            }

        }

        public MultiValueDialog()
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
            dataSource.DataSource = dataSet1;

            //AutoGenerateColumns ?]?w?n?? DataSource ?e??
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = dataSource;


            LookupDataset = dataSet1;
            
		}

        //????
		protected virtual void ProcessOK()
		{

            if (BindDataSource.Current != null)
            {
                //LookupValues = new object[dataGridView1.Columns.Count];


                //_selectId = (string)((DataRowView)((BindingSource)BindDataSource).Current)[0];
                //_selectName = (string)((DataRowView)((BindingSource)BindDataSource).Current)[1];


                //for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
                //{
                //    LookupValues[i] = ((DataRowView)((BindingSource)BindDataSource).Current)[i];
                //}

                LookupValues = new object[dataGridView1.SelectedRows.Count];
                //?????????W..????????
                DataGridViewRow row = null;
               int iCount = 0;
                for (int i = dataGridView1.SelectedRows.Count - 1; i >= 0; i--)
                
               // for (int i = 0; i <= dataGridView1.SelectedRows.Count - 1; i++)
                {
                    row = dataGridView1.SelectedRows[i];
                    LookupValues[iCount] = row.Cells[KeyFieldName].Value.ToString();
                    iCount++;
                }

            }
		}

        //?W?[?d??????
        protected virtual void InitFields(ComboBox cb)
        {
            cbFields = cb;
            //cb.Items.Add("??????(TABLE_NAME)");
            //cb.Items.Add("????(FIELD_NAME)");
            for (int i = 0; i <= FieldNames.Length - 1; i++)
            {
                //CreateColumn(FieldNames[i], Captions[i]);
                cb.Items.Add(string.Format("{0}({1})", Captions[i], FieldNames[i]));
            }


            cb.SelectedIndex = 0;
        }


        private void CreateColumn(string FieldName, string Caption)
        {
            DataGridViewTextBoxColumn textColumn = new DataGridViewTextBoxColumn();

            textColumn.DataPropertyName = FieldName;

            textColumn.Name = FieldName;
            textColumn.HeaderText = Caption;

            dataGridView1.Columns.Add(textColumn);
        }

        //
        protected  void DoQuery(DataTable dt)
        {
            DataView dv = new DataView(dt);

            if (tbExpression.Text.Trim() == "")
            {
                MessageBox.Show("?????J?d?????????I", "?H??????");
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

                MessageBox.Show("???@?X?{???~?G" + err.Message, "?H??????", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

            for (int i = 0; i <= FieldNames.Length - 1; i++)
            {
                CreateColumn(FieldNames[i], Captions[i]);

            }



            ViewData();

            InitializeForm(bindingSource1);

            //?S???? ???n?? baseform ?? TabIndex
            dataGridView1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DoQuery(LookupDataset);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Selected = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i <= dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Selected = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            for (int i = dataGridView1.CurrentRow.Index; i <= dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Selected = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            for (int i = dataGridView1.CurrentRow.Index; i <= dataGridView1.Rows.Count - 1; i++)
            {
                dataGridView1.Rows[i].Selected = false;
            }
        }

        public void ViewData()
        {


            try
            {
                this.Cursor = Cursors.WaitCursor;


                //?u??????
                if (!string.IsNullOrEmpty(SqlScript))
                {

                    if (LookUpConnection == null)
                    {
                        dataSet1 = GetData(SqlScript);
                    }
                    else
                    {
                        dataSet1 = GetData(LookUpConnection, SqlScript);
                    }
                    return;
                }


                if (SourceDataTable != null)
                {
                    dataSet1 = SourceDataTable;
                }


            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        //?p?????}??
        //?W?L1?U????????,?????????L????//???O?]?? top 10000
        public DataTable GetCOLDEF()
        {
            SqlConnection connection = globals.shipConnection;
            string sql = "SELECT TABLE_NAME,FIELD_NAME,CAPTION FROM COLDEF";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@TABLE_NAME", TABLE_NAME));
            //command.Parameters.Add(new SqlParameter("@FIELD_TYPE", FIELD_TYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "COLDEF");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["COLDEF"];
        }

        public DataTable GetData(string SqlScript)
        {
            SqlConnection connection = globals.shipConnection;
            string sql = SqlScript;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@TABLE_NAME", TABLE_NAME));
            //command.Parameters.Add(new SqlParameter("@FIELD_TYPE", FIELD_TYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Data");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["Data"];
        }


        public DataTable GetData(SqlConnection connection, string SqlScript)
        {
            SqlConnection connections = globals.shipConnection;
            string sql = SqlScript;
            SqlCommand command = new SqlCommand(sql, connections);
            command.CommandType = CommandType.Text;
            //command.Parameters.Add(new SqlParameter("@TABLE_NAME", TABLE_NAME));
            //command.Parameters.Add(new SqlParameter("@FIELD_TYPE", FIELD_TYPE));
            SqlDataAdapter da = new SqlDataAdapter(command);
            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "Data");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["Data"];
        }
       
	}
}
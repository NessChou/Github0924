using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
//�ϥνd��
//�ǤJ��
//���         string[] FieldNames
//��줤��W�� string[] Captions
//SQL�y�k      SqlScript
//�Ǧ^�� 


namespace ACME
{
    public partial class CHOLookup5 : ACME.LookupDialog
    {
        string strCn = "Data Source=10.10.1.40;Initial Catalog=CHIComp16;Persist Security Info=True;User ID=webstock;Password=@cmewebstock";
        public CHOLookup5()
        {
            InitializeComponent();
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


        private object[]  fLookupValues;

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


        protected override void ProcessOK()
        {
            if (BindDataSource.Current != null)
            {
                LookupValues = new object[dataGridView1.Columns.Count];
                
                
                //_selectId = (string)((DataRowView)((BindingSource)BindDataSource).Current)[0];
                //_selectName = (string)((DataRowView)((BindingSource)BindDataSource).Current)[1];


                for (int i = 0; i <= dataGridView1.Columns.Count - 1; i++)
                {
                    LookupValues[i] = ((DataRowView)((BindingSource)BindDataSource).Current)[i];
                }
            }
        }

        protected override void InitFields(ComboBox cb)
        {
            base.InitFields(cb);

            //cb.Items.Add("��ƪ�(TABLE_NAME)");
            //cb.Items.Add("���(FIELD_NAME)");
            for (int i = 0; i <= FieldNames.Length - 1; i++)
            {
                //CreateColumn(FieldNames[i], Captions[i]);
                cb.Items.Add(string.Format("{0}({1})",Captions[i],FieldNames[i]));
            }


            cb.SelectedIndex = 0;
           
        }




        private void CreateColumn(string FieldName,string Caption)
        {
            DataGridViewTextBoxColumn textColumn = new DataGridViewTextBoxColumn();

            textColumn.DataPropertyName = FieldName;

            textColumn.Name = FieldName;
            textColumn.HeaderText = Caption;
            textColumn.Width = 150;
            dataGridView1.Columns.Add(textColumn);
        }



        private void EmpLookup_Load(object sender, EventArgs e)
        {

            for (int i = 0; i <= FieldNames.Length - 1; i++)
            {
                CreateColumn(FieldNames[i], Captions[i]);

            }



            ViewData();

            InitializeForm(bindingSource1);
            
            //�S���� �ݭn�� baseform �� TabIndex
            dataGridView1.Focus();
        }

        public void ViewData()
        {


            try
            {
                this.Cursor = Cursors.WaitCursor;


                //�u������
                if (!string.IsNullOrEmpty(SqlScript))
                {
                    dataSet1 = GetData(SqlScript);
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


        protected override void InitializeForm(BindingSource dataSource)
        {
            base.InitializeForm(dataSource);

            dataSource.DataSource = dataSet1;

            //AutoGenerateColumns �]�w�n�� DataSource �e��
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = dataSource;
            

            LookupDataset = dataSet1;
        }

      

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            btnOK_Click(sender, EventArgs.Empty);
        }

        //�p�ɮ׶}��
        //�W�L1�U����Ʈ�,�ШϥΨ�L�覡//�άO�]�� top 10000
        public  DataTable GetCOLDEF()
        {
   
            SqlConnection connection = new SqlConnection(strCn);

            string sql = "SELECT TABLE_NAME,FIELD_NAME,CAPTION FROM COLDEF";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            
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
        
            SqlConnection connection = new SqlConnection(strCn);
            string sql = SqlScript;
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
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


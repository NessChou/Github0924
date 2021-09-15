using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using Microsoft.VisualBasic.Devices;
using System.Data.SqlClient;

//
//�d�ߤ���
using DragD.QuickWhereComponent;

namespace ACME
{
    public partial class fmBase5 : Form
    {
        string strCn = "Data Source=acmesap;Initial Catalog=acmesqlsp;Persist Security Info=True;User ID=sapdbo;Password=@rmas";
        public fmBase5()
        {
            this.FAllowAddNew = true;
            this.FAllowEdit = true;
            this.FAllowPrint = false;
            //
            this.FAllowDelete = true;

            
            InitializeComponent();
        }

        //���\��s
        private bool FAllowAddNew;
        //���\�ק�
        private bool FAllowEdit;

        //
        private bool FAllowDelete;

        //���\�C�L���
        private bool FAllowPrint;
        //�s������
        private SqlConnection FConnection;
        //�D���
        private string FID;
        private string STOPID;
        private string STOPID2;
        //��ƪ�D�䪺���W��
        private string FIDFieldName;
        //��ƪ��A
        private string FStatus;
        //��ƪ�W��
        private string FTableName;
        //
       
        private BindingSource FBindingSource;
        //SQL �y�k
        private string FSelectSQL;

        public  string kyes;
        //��Ʀr��

        private DataTable dtColDef;

        //�d��
        private QuickWhere QW;

        //Where ����
        private string QuickWhere_SqlScript;

        //��Ƶ��� //�`���Ʒ|�H�O���W�[�ӧ���
        //���D�N��ɪ��d�߰O���Ȧs
        //�Ҧ�1-�Y��
        //�Ҧ�2-���d�߼Ȧs��->�u�s ID ���
        //�p�G�� QuickWhere_SqlScript ��2
        //�w�]�� 1
        private int  FRecordCount;

        private int FRecordNo;

        //Master Table
        private DataTable FMasterTable;
        // private DataTableCollection FDetailTables;
        private DataTable[] FDetailTables;
        private BindingSource[] FDetailBindingSources;

        private DataTableCollection CloneTables;

        public DataTable MasterTable
        {
            get
            {
                return this.FMasterTable;
            }
            set
            {
                this.FMasterTable = value;
            }
        }
     

        public DataTable[] DetailTables
        {
            get
            {
                return this.FDetailTables;
            }
            set
            {
                this.FDetailTables = value;
            }
        }

        public BindingSource[] DetailBindingSources
        {
            get
            {
                return this.FDetailBindingSources;
            }
            set
            {
                this.FDetailBindingSources = value;
            }
        }
        public virtual void AfterAddNew()
        {
        }

        public virtual void AfterCancelEdit()
        {
        }
        public virtual void UpdareEnd()
        {
        }
        public virtual void BBowhsload()
        {
        }
        public virtual void EndEdit()
        {
        }
        public virtual void EndEdit2()
        {
        }
        public virtual void AfterEdit()
        {
        }

        public virtual void query()
        {
        }
        public virtual void Query()
        {
        }
 
   
   
        public virtual void STOP()
        {
        }
        public virtual void STOP2()
        {
        }
        public virtual void AfterCopy()
        {
        }
        public virtual void AfterCopy2()
        {
        }
        public virtual void AfterEndEdit()
        {
         
            GetMaxorNext();

        }

       
        public virtual void AfterLoad()
        {
        }

        public virtual void AfterScroll()
        {
        }

        public virtual bool BeforeAddNew()
        {
            return true;
        }

        public virtual bool BeforeCancelEdit()
        {
            return true;
        }


        public virtual bool BeforeDelete()
        {
            return true;
        }

        public virtual bool BeforeEndEdit()
        {
            return true;
        }
        public virtual void AfterDelete()
        {
        }
        
        public virtual void BeforeLoad()
        {
        }
        public virtual void SAVE()
        {
        }

        public virtual bool BeforeScroll()
        {
            return true;
        }

        public virtual void FillData()
        {

        }
      
        public virtual void SetControls()
        {
        }

        public virtual void SetDefaultValue()
        {
        }
      
        public virtual bool UpdateData()
        {
            return true;
        }
      
      
        public virtual void DoPrint()
        {
        }

        public virtual void SetConnection()
        {
            MyConnection = globals.Connection;
        }

        public virtual void SetInit()
        {
            
        }


        //�ץ�
        //if (!string.IsNullOrEmpty(NewID) )
        private void bnNext_Click(object sender, EventArgs e)
        {
            if (this.BeforeScroll() && (MyID != null))
            {
                string sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + " WHERE " + this.MyIDFieldName + " > '" + this.MyID + "' AND BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'  ORDER BY SHIPPINGCODE ";
               // string sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + " WHERE " + this.MyIDFieldName + " > '" + this.MyID + "'";
                
                //

                if (!string.IsNullOrEmpty(QuickWhere_SqlScript))
                {
                    sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + "  WHERE " + this.MyIDFieldName + " > '" + this.MyID + "' AND BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'  "
                        + " AND " + QuickWhere_SqlScript + "ORDER BY SHIPPINGCODE";
                }

                //SelectSQL = sSQL;

                SqlCommand cmdSQL = new SqlCommand();
                cmdSQL.CommandText = sSQL;
                cmdSQL.Connection = this.MyConnection;
                this.MyConnection.Open();
                string NewID = Convert.ToString(cmdSQL.ExecuteScalar());
                this.MyConnection.Close();
                if (!string.IsNullOrEmpty(NewID) )
                {
                    this.MyID = NewID;
                    this.BeforeLoad();
                    this.FillData();
                    this.AfterLoad();
                }
                this.MyTableStatus = "0";
                SL_Status.Text = "�s��";

                if (!string.IsNullOrEmpty(QuickWhere_SqlScript))
                {
                    if (RecordNo == RecordCount)
                    {
                    }
                    else
                    {
                        RecordNo++;
                    }
                }

                ShowRecordCount(RecordNo, RecordCount);

                this.AfterScroll();
            }
            this.SetButtons();

        }

        //�s�W
        private void bnAddNew_Click(object sender, EventArgs e)
        {
      
                this.MyBS.AddNew();
                this.MyTableStatus = "1";
                SL_Status.Text = "�s�W�@";
                this.SetControls();
                this.SetDefaultValue();
        

            
            this.SetButtons();
            SetControlEnabled(Controls, true);
            AfterAddNew();
            Copy2.Enabled = false;

            SearchTextBox.Visible = false;
            bnSearch.Visible = false;
            
        }


        private void SetButtons()
        {

            if (!string.IsNullOrEmpty(QuickWhere_SqlScript))
            {
                this.bnFirst.Enabled = (this.FStatus == "0") && (RecordNo != 1);
                this.bnPrevious.Enabled = (this.FStatus == "0") && (RecordNo != 1);
                this.bnNext.Enabled = (this.FStatus == "0") && (RecordNo != RecordCount);
                this.bnLast.Enabled = (this.FStatus == "0") && (RecordNo != RecordCount);
            }
            else
            {
                this.bnFirst.Enabled = (this.FStatus == "0") ;
                this.bnPrevious.Enabled = (this.FStatus == "0") ;
                this.bnNext.Enabled = (this.FStatus == "0") ;
                this.bnLast.Enabled = (this.FStatus == "0") ;
            }

            this.bnAddNew.Enabled = (this.FStatus == "0") &this.AllowAddNew;
            //

            this.bnDelete.Enabled = (this.FStatus == "0") & this.AllowDelete;

            this.bnEdit.Enabled =  (this.FStatus == "0") & this.MyID != null & this.AllowEdit;

            this.bnEndEdit.Enabled = (this.FStatus == "1") | (this.FStatus == "2") | (this.FStatus == "3") | (this.FStatus == "9");

            this.bnCancelEdit.Enabled = (this.FStatus == "1") | (this.FStatus == "2") | (this.FStatus == "3") | (this.FStatus == "9");



           // this.bnSearch.Enabled = (this.FStatus == "0") & this.SearchTextBox.Text != "";
            this.bnPrint.Enabled = (this.FStatus == "0") &this.AllowPrint;
            this.SearchTextBox.ReadOnly = (this.FStatus != "0");
        

            //�d��
            bnQuery.Enabled = (this.FStatus == "0");





        }


        // Properties
        public bool AllowAddNew
        {
            get
            {
                return this.FAllowAddNew;
            }
            set
            {
                this.FAllowAddNew = value;
                this.SetButtons();
            }
        }

        //
        public bool AllowDelete
        {
            get
            {
                return this.FAllowDelete;
            }
            set
            {
                this.FAllowDelete = value;
                this.SetButtons();
            }
        }

        public bool AllowEdit
        {
            get
            {
                return this.FAllowEdit;
            }
            set
            {
                this.FAllowEdit = value;
                this.SetButtons();
            }
        }

        public bool AllowPrint
        {
            get
            {
                return this.FAllowPrint;
            }
            set
            {
                this.FAllowPrint = value;
                this.SetButtons();
            }
        }

        public BindingSource MyBS
        {
            get
            {
                return this.FBindingSource;
            }
            set
            {
                this.FBindingSource = value;
                this.BaseBindingNavigator.BindingSource = value;
            }
        }

        public SqlConnection MyConnection
        {
            get
            {
                return this.FConnection;
            }
            set
            {
                this.FConnection = value;
            }
        }

        public string MyID
        {
            get
            {
                return this.FID;
            }
            set
            {
                this.FID = value;
            }
        }
        public string SSTOPID
        {
            get
            {
                return this.STOPID;
            }
            set
            {
                this.STOPID = value;
            }
        }
        public string SSTOPID2
        {
            get
            {
                return this.STOPID2;
            }
            set
            {
                this.STOPID2 = value;
            }
        }
        public string SelectSQL
        {
            get
            {
                return this.FSelectSQL;
            }
            set
            {
                this.FSelectSQL = value;
            }
        }


        public string MyIDFieldName
        {
            get
            {
                return this.FIDFieldName;
            }
            set
            {
                this.FIDFieldName = value;
            }
        }

        public string MyTableName
        {
            get
            {
                return this.FTableName;
            }
            set
            {
                this.FTableName = value;
            }
        }

        public string MyTableStatus
        {
            get
            {
                return this.FStatus;
            }
            set
            {
                this.FStatus = value;
            }
        }


        public int RecordCount
        {
            get
            {
                return this.FRecordCount;
            }
            set
            {
                this.FRecordCount = value;
            }
        }


        public int RecordNo
        {
            get
            {
                return this.FRecordNo;
            }
            set
            {
                this.FRecordNo = value;
            }
        }



        private void bnEdit_Click(object sender, EventArgs e)
        {
            SSTOPID2 = "0";
            STOP2();
            if (SSTOPID2 == "1")
            {
                return;
            }
                this.MyTableStatus = "2";
                SL_Status.Text = "�ק�";

            //20140415
                this.FillData();


                this.SetControls();
 
            this.SetButtons();
            SetControlEnabled(Controls, true);

            this.AfterEdit();
      
            Copy2.Enabled = false;

            SearchTextBox.Visible = false;
            bnSearch.Visible = false;

        }

        private void bnEndEdit_Click(object sender, EventArgs e)
        {
           
            Copy2.Enabled = true;

            //if (MyTableName == "Shipping_main")
            //{
            if (MyTableName != "SATT" && MyTableName != "Account_LC")
            {
                bnSearch.Visible = true;
                SearchTextBox.Visible = true;
            }
           // }


            //�W�[�d�߼Ҧ�
            if (MyTableStatus == "9")
            {
                QW.Clear();
                SetControlWhere(Controls);
                QuickWhere_SqlScript = QW.GetSql();


                this.MyBS.CancelEdit();


                if (this.BeforeScroll() && (MyID != null))
                {
                    string sSQL = string.Empty;
                    if (!string.IsNullOrEmpty(QuickWhere_SqlScript))
                    {

                        sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + " WHERE " + QuickWhere_SqlScript + " AND BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'  ORDER BY SHIPPINGCODE ";
                        SqlCommand cmdSQL = new SqlCommand();
                        cmdSQL.CommandText = sSQL;
                        cmdSQL.Connection = this.MyConnection;
                        this.MyConnection.Open();
                        string NewID = Convert.ToString(cmdSQL.ExecuteScalar());
                        this.MyConnection.Close();
                        if (!string.IsNullOrEmpty(NewID))
                        {
                            this.MyID = NewID;
                            this.BeforeLoad();
                            this.FillData();
                            this.AfterLoad();
                        }

                        this.AfterScroll();

                        RecordCount = GetRecordCount(QuickWhere_SqlScript);
                        if (RecordCount == 0)
                        {
                            RecordNo = 0;
                            ShowRecordCount(RecordNo, RecordCount);
                            MessageBox.Show("�d�L�O��");
                        }
                        else
                        {
                            RecordNo = 1;
                            ShowRecordCount(RecordNo, RecordCount);
                        }
                    }
                    else
                    {
                        GetMaxRecord();
                        RecordNo = 0;
                        RecordCount = 0;
                        ShowRecordCount(RecordNo, RecordCount);
                    }



                }

                this.MyTableStatus = "0";
                SL_Status.Text = "�s��";
                this.SetControls();
                this.SetButtons();
      
                SetControlEnabled(Controls, false);
                EndEdit();
                return;
               


            }

            else 
            {
                SSTOPID = "0";
                STOP();
                if (SSTOPID == "1")
                {
                    return;
                }
            }


            this.BBowhsload();
            if (this.BeforeEndEdit())
            {
                string PrevTableStatus = this.MyTableStatus;
                try
                {
                    if (MyTableStatus == "3")
                    {

                    }



                    if (this.UpdateData())
                    {


                        this.AfterEndEdit();
                        this.MyTableStatus = "0";
                        SL_Status.Text = "�s��";
                        this.SetControls();



                    }
                }
                catch (Exception exception1)
                {
                    Exception ex = exception1;
                    MessageBox.Show(ex.Message, "�ާ@���~", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    this.MyTableStatus = PrevTableStatus;

                }
            }
            this.SetButtons();
            SetControlEnabled(Controls, false);
            this.EndEdit();
            this.EndEdit2();
        }

        private void bnCancelEdit_Click(object sender, EventArgs e)
        {
           
            if (this.BeforeCancelEdit())
            {
                this.MyBS.CancelEdit();
                this.MyTableStatus = "0";
                SL_Status.Text = "�s��";
                this.SetControls();
            

                this.SetButtons();
                SetControlEnabled(Controls, false);
                this.AfterCancelEdit();
                Copy2.Enabled = true;


                //if (MyTableName == "Shipping_main")
                //{
                if (MyTableName != "SATT" && MyTableName != "Account_LC")
                {
                    bnSearch.Visible = true;
                    SearchTextBox.Visible = true;
                }
               // }

            }

          
   
        }

        private void bnSearch_Click(object sender, EventArgs e)
        {
            this.MyID = this.SearchTextBox.Text.Trim();
            if (!String.IsNullOrEmpty(this.MyID))
            {

                string sSQL = string.Empty;
                if (FieldComboBox.Items.Count > 0)
                {
                    string QueryFieldName = Convert.ToString(dtColDef.Rows[FieldComboBox.SelectedIndex]["FIELD_NAME"]);


                    sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + " WHERE " + QueryFieldName + " LIKE '" + this.MyID + "%' AND BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'   ";

                }
                else
                {
                    sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + " WHERE " + this.MyIDFieldName + " LIKE '" + this.MyID + "%' AND BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'  ";
                }

                SqlCommand cmdSQL = new SqlCommand();
                cmdSQL.CommandText =sSQL;
                cmdSQL.Connection = this.MyConnection;
                this.MyConnection.Open();
                string NewID = Convert.ToString(cmdSQL.ExecuteScalar());
                this.MyConnection.Close();
                if (!String.IsNullOrEmpty(NewID))
                {
                    this.MyID = NewID;
                    this.BeforeLoad();
                    this.FillData();
                    this.AfterLoad();
                }
                this.MyTableStatus = "0";
                SL_Status.Text = "�s��";
                this.AfterScroll();
                QuickWhere_SqlScript = "";
                this.SetButtons();
            }

            SearchTextBox.Text = "";

        }

        private void bnPrint_Click(object sender, EventArgs e)
        {
            DoPrint();
        }

        private void bnFirst_Click(object sender, EventArgs e)
        {
            if (this.BeforeScroll())
            {
                string sSQL = "SELECT TOP 1 (" + this.MyIDFieldName + ") AS ID FROM " + this.MyTableName;

               // string sSQL = "SELECT MIN(" + this.MyIDFieldName + ") AS ID FROM " + this.MyTableName;

                if (!string.IsNullOrEmpty(QuickWhere_SqlScript))
                {
                    sSQL = sSQL + " WHERE BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'  AND " + QuickWhere_SqlScript;
                }
                sSQL = sSQL + " ORDER BY SHIPPINGCODE ";

                //SelectSQL = sSQL;

                SqlCommand cmdSQL = new SqlCommand();
                cmdSQL.CommandText = sSQL;
                cmdSQL.Connection = this.MyConnection;
                this.MyConnection.Open();
                string NewID = cmdSQL.ExecuteScalar().ToString();
                this.MyConnection.Close();
                if (NewID!=null)
                {
                    this.MyID = NewID;
                    this.BeforeLoad();
                    this.FillData();
                    this.AfterLoad();
                }
                this.MyTableStatus = "0";
                SL_Status.Text = "�s��";
                RecordNo = 1;

                ShowRecordCount(RecordNo, RecordCount);
                
                this.AfterScroll();
            }
            this.SetButtons();

        }

        private void bnPrevious_Click(object sender, EventArgs e)
        {
            if (this.BeforeScroll())
            {
                string sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + " WHERE " + this.MyIDFieldName + " < '" + this.MyID + "' AND BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'   ORDER BY SHIPPINGCODE DESC";


                if (!string.IsNullOrEmpty(QuickWhere_SqlScript))
                {
                    sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + " WHERE " + this.MyIDFieldName + " < '" + this.MyID + "'"
                        + " AND " + QuickWhere_SqlScript + " AND BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'  ORDER BY SHIPPINGCODE  DESC";
                }

            //    string sSQL = "SELECT TOP 1 (" + this.MyIDFieldName + ") AS ID FROM " + this.MyTableName + " ORDER BY REPLACE((" + this.MyIDFieldName + "),'DRS','') DESC ";

                //SelectSQL = sSQL;

                SqlCommand cmdSQL = new SqlCommand();
                cmdSQL.CommandText =sSQL;
                cmdSQL.Connection = this.MyConnection;
                this.MyConnection.Open();
                string NewID = Convert.ToString(cmdSQL.ExecuteScalar());
                this.MyConnection.Close();
                if (!string.IsNullOrEmpty(NewID))
                {
                    this.MyID = NewID;
                    this.BeforeLoad();
                    this.FillData();
                    this.AfterLoad();
                }
                this.MyTableStatus = "0";
                SL_Status.Text = "�s��";


                if (!string.IsNullOrEmpty(QuickWhere_SqlScript))
                {
                    if (RecordNo == 1)
                    {
                    }
                    else
                    {
                        RecordNo--;
                    }
                }
                ShowRecordCount(RecordNo, RecordCount);

                this.AfterScroll();
            }
            this.SetButtons();

        }

        private void bnLast_Click(object sender, EventArgs e)
        {
            if (this.BeforeScroll())
            {
                string sSQL = "SELECT TOP 1 (" + this.MyIDFieldName + ") AS ID FROM " + this.MyTableName ;

                //string sSQL = "SELECT TOP 1 MAX(" + this.MyIDFieldName + ") AS ID FROM " + this.MyTableName;

                if (!string.IsNullOrEmpty(QuickWhere_SqlScript))
                {
                    sSQL = sSQL + " WHERE  BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'   AND " + QuickWhere_SqlScript;
                }

                sSQL = sSQL + " ORDER BY SHIPPINGCODE DESC ";

                //SelectSQL = sSQL;

                SqlCommand cmdSQL = new SqlCommand();
                cmdSQL.CommandText = sSQL;
                cmdSQL.Connection = this.MyConnection;
                this.MyConnection.Open();
                string NewID = cmdSQL.ExecuteScalar().ToString();
                this.MyConnection.Close();
                if (NewID != null)
                {
                    this.MyID = NewID;
                    this.BeforeLoad();
                    this.FillData();
                    this.AfterLoad();
                }
                this.MyTableStatus = "0";
                SL_Status.Text = "�s��";

                RecordNo = RecordCount;

                ShowRecordCount(RecordNo, RecordCount);

                this.AfterScroll();
            }
            this.SetButtons();

        }

        protected void fmBase_Load(object sender, EventArgs e)
        {

            if (!DesignMode)
            {
                //�ܦ��̤j�Ӥ��|���Ż�
                this.WindowState = FormWindowState.Maximized;
                

                SetConnection();
                SetInit();


                if (MyConnection == null)
                {

                    MessageBox.Show("�г]�w MyConnection!!!");
                    Close();
                    return;
                }

                GetMaxRecord();

                //if (!string.IsNullOrEmpty(MyTableName))
                //{
                //    //try
                //    //{
                //    //    dtColDef = GetCOLDEF(MyTableName, "VARCHAR");
                //    //}
                //    //catch
                //    //{ 
                    
                //    //}

                //    try
                //    {
                //        SetFieldComboBox();
                //    }
                //    catch
                //    { 
                    
                //    }

                //}
                this.MyTableStatus = "0";
                SL_Status.Text = "�s��";
                
                this.Left = 0;
                this.Top = 0;


                //�]�w�v��

                SetAuthority();

                if (MyTableName == "Shipping_main" || MyTableName == "Shipping_CAR")
                {
                    SAVEButton.Visible = true;
                    //bnSearch.Visible = true;
                    //SearchTextBox.Visible = true;
      
                }

                //Shipping_main
                if (MyTableName != "SATT" && MyTableName != "Account_LC")
                {
                    bnSearch.Visible = true;
                    SearchTextBox.Visible = true;
                }
             

                this.SetButtons();
                this.SetControls();

                SetControlEnabled(Controls, false);
                SetControlEvent(Controls);
            
                //�d�ߤ����l��
                QW = new QuickWhere();
                //�[�J�r��
                QuickWhere.SetGenerals('\'', '#', '?', '%', ',', '\\', "@@");


                if (MyTableName == "SATT" || MyTableName == "Shipping_main")
                {
                    if (globals.DBNAME == "�F�ͥ�")
                    {
                        if (globals.GroupID.ToString().Trim() == "SA")
                        {
                            bnAddNew.Visible = true;
                            if (MyTableName == "Shipping_main")
                            {
                                bnEdit.Visible = true;
                            }
                        }
                    }
                }
            }

        }

        private void SetFieldComboBox()
        {
            for (int i = 0; i <= dtColDef.Rows.Count - 1; i++)
            {

                FieldComboBox.Items.Add(Convert.ToString(dtColDef.Rows[i]["Caption"]));
            }
        
        }

        public DataTable GetOrdr3(string username)
        {
            SqlConnection connection = new SqlConnection(strCn);

            string sql = "select category from [right] where username=@username";
            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@username", username));
            SqlDataAdapter da = new SqlDataAdapter(command);
            DataSet ds = new DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "right");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["right"];
        }
        //�]�w�v��
        //APPEND, EDIT, DEL, REPORT
        private void SetAuthority()
        {
            System.Data.DataTable dt3 = GetOrdr3(fmLogin.LoginID.ToString());
            string GG1 = "";
            if (dt3.Rows.Count > 0)
            {
                DataRow drw = dt3.Rows[0];
                GG1 = drw["category"].ToString();
            }

            string aa = GG1.Trim();
            string[] arrurl = aa.Split(new Char[] { ',' });
            StringBuilder sb = new StringBuilder();
            foreach (string i in arrurl)
            {
                sb.Append("'" + i + "',");
            }
            sb.Remove(sb.Length - 1, 1);

            int GG2 = aa.IndexOf(",");
            if (GG2 != -1)
            {
                globals.GroupID = aa.Substring(0, GG2);
            }
            else
            {
                globals.GroupID = aa;
            }

            DataTable dtUserMenu = GetUSERMENUS(sb.ToString(), this.Name);

            string APPEND =Convert.ToString(dtUserMenu.Rows[0]["APPEND"]);
            string EDIT = Convert.ToString(dtUserMenu.Rows[0]["EDIT"]);
            string DEL = Convert.ToString(dtUserMenu.Rows[0]["DEL"]);
            string REPORT = Convert.ToString(dtUserMenu.Rows[0]["REPORT"]);
            string SYSFLAG = Convert.ToString(dtUserMenu.Rows[0]["SYSFLAG"]);
            if (APPEND != "Y")
            {
                bnAddNew.Visible = false;
            }

            if (EDIT != "Y")
            {
                bnEdit.Visible = false;
            }

            if (DEL != "Y")
            {
                bnDelete.Visible = false;
            }

            if (REPORT != "Y")
            {
                Copy2.Visible = false;
                bnPrint.Visible = false;
            }
            if (SYSFLAG != "Y")
            {
                bnQuery.Visible = false;
   
            }
        }


 

        private void GetMaxRecord()
        {
          //  string sSQL = "SELECT MAX(" + this.MyIDFieldName + ") AS ID FROM " + this.MyTableName;
            string sSQL = "SELECT TOP 1 (" + this.MyIDFieldName + ") AS ID FROM " + this.MyTableName + " WHERE BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'   ORDER BY SHIPPINGCODE DESC ";
            SqlCommand cmdSQL = new SqlCommand();
            cmdSQL.CommandText =sSQL;
            cmdSQL.Connection = MyConnection;
            this.MyConnection.Open();

            string NewID = string.Empty;

            try
            {
                NewID = cmdSQL.ExecuteScalar().ToString();
            }
            catch
            {
                NewID = null;
            }

            finally
            {
                this.MyConnection.Close();
            }
            if (NewID != null)
            {
                this.MyID = NewID;
                this.FillData();
            }
        
        }




        //�R������̤j���ΤU�@��
        private void GetMaxorNext()
        {
            if (MyTableStatus == "3")
            {

                string sSQL = "SELECT TOP 1 " + this.MyIDFieldName + " FROM " + this.MyTableName + " WHERE " + this.MyIDFieldName + " > '" + this.MyID + "' AND BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D'   ";
                SqlCommand cmdSQL = new SqlCommand();
                cmdSQL.CommandText =sSQL;
                cmdSQL.Connection = this.MyConnection;
                this.MyConnection.Open();

                string NewID = string.Empty;
                try
                {
                    NewID = cmdSQL.ExecuteScalar().ToString();
                }
                catch
                {
                    sSQL = "SELECT MAX(" + this.MyIDFieldName + ") AS ID FROM " + this.MyTableName;
                    cmdSQL.CommandText = Convert.ToString(sSQL);
                    NewID = cmdSQL.ExecuteScalar().ToString();
                }
                this.MyConnection.Close();


                if (NewID != null)
                {
                    this.MyID = NewID;
                    this.FillData();
                }
            }
        }

        public  int GetRecordCount()
        {
           
            string sql = "SELECT COUNT(*) FROM " + this.MyTableName ;
            SqlCommand command = new SqlCommand(sql, this.MyConnection);
            command.CommandType = CommandType.Text;
            try
            {
                MyConnection.Open();
                return (Int32)command.ExecuteScalar();
            }
            finally
            {
                MyConnection.Close();
            }
        }


        public int GetRecordCount(string WhereStr)
        {

            string sql = "SELECT COUNT(*) FROM " + this.MyTableName + " WHERE  BoardCountNo ='�T��' AND SUBSTRING(SHIPPINGCODE,1,2)<> 'SI'  AND SUBSTRING(SHIPPINGCODE,14,1)<>'D' AND " + WhereStr;
            SqlCommand command = new SqlCommand(sql, this.MyConnection);
            command.CommandType = CommandType.Text;
            try
            {
                MyConnection.Open();
                return (Int32)command.ExecuteScalar();
            }
            finally
            {
                MyConnection.Close();
            }
        }

        private void fmBase_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (this.MyTableStatus != "0")
            {
                if (MessageBox.Show("��Ʃ|���x�s, ���n���}���@�~��?", "��ƥ��x�s", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == DialogResult.No)
                {
                    e.Cancel = true;
                }
                else
                {
                    e.Cancel = false;
                }
            }

        }


        private void FBindingSource_PositionChanged(object sender, EventArgs e)
        {
            this.SetButtons();
        }

        private void bnDelete_Click(object sender, EventArgs e)
        {
            if (this.BeforeDelete() & (this.MyID != null))
            {
                this.MyTableStatus = "3";
                SL_Status.Text = "�R��";
                this.SetControls();
                this.AfterDelete();
            }
            this.SetButtons();

        }


        //�C�Ӧr��J��N����...
        private void SearchTextBox_TextChanged(object sender, EventArgs e)
        {
            //this.bnSearch.Enabled = (this.FStatus== "0") & (this.SearchTextBox.Text != "");

        }


        //Enabled
        public void SetControlEnabled(System.Windows.Forms.Control.ControlCollection originalControls, bool EnabledFlag)
        {
           // Control aControl;
            bool anti;
            if (EnabledFlag == true)
            {
                anti = false;
            }
            else
            {
                anti = true;
            }
            for (int i = 0; i <= originalControls.Count - 1; i++)
            {
                if (originalControls[i].Controls.Count > 0)
                {

                    SetControlEnabled(originalControls[i].Controls, EnabledFlag);
                }

                if (originalControls[i] is TextBox)
                {

          
                    TextBox aTextBox = (TextBox)originalControls[i];

                    //aTextBox.Enabled = EnabledFlag;
                    

                    aTextBox.ReadOnly = anti;
                    //�ק� Enabled ���C��
                    if (EnabledFlag)
                    {
                        aTextBox.BackColor = Color.White;
                        aTextBox.ForeColor = Color.Black;
                   
                    }
                    else
                    {
                        aTextBox.BackColor = Color.White;
                        aTextBox.ForeColor = Color.Black;
                        
                      //  MessageBox.Show("");
                        
                    }
                   // aTextBox.ReadOnly = ! Enabled;
                }

       
                if (originalControls[i] is CheckBox)
                {


                    CheckBox aTextBox = (CheckBox)originalControls[i];

                    //aTextBox.Enabled = EnabledFlag;


                    aTextBox.Enabled = EnabledFlag;

                }
                if (originalControls[i] is Button)
                {


                    Button aTextBox = (Button)originalControls[i];

                    //aTextBox.Enabled = EnabledFlag;


                    aTextBox.Enabled = EnabledFlag;
                    //�ק� Enabled ���C��

                    // aTextBox.ReadOnly = ! Enabled;
                }
                if (originalControls[i] is ComboBox)
                {


                    ComboBox aTextBox = (ComboBox)originalControls[i];
                    //DropDownList �~�|�C���ܹ�
                    aTextBox.DropDownStyle = ComboBoxStyle.DropDownList;
                    aTextBox.Enabled = EnabledFlag;
                    //DropDownList �~�|�C���ܹ�
                 //   aTextBox.r = anti;
                    //�ק� Enabled ���C��
                    if (EnabledFlag)
                    {
                        aTextBox.BackColor = Color.White;
                        aTextBox.ForeColor = Color.Black;
                    }
                    else
                    {
                        aTextBox.BackColor = Color.White;
                        aTextBox.ForeColor = Color.Black;

                    }
                
                }

                try
                {
                    if (originalControls[i] is DataGridView)
                    {


                        DataGridView DataGridView = (DataGridView)originalControls[i];

                     //   DataGridView.Enabled = EnabledFlag;
                        DataGridView.ReadOnly = anti;
                    }

            
           
                }
                catch { 
                }
        
            }
            SearchTextBox.Enabled = true;
            SearchTextBox.ReadOnly = false;
        }
        //Enabled
        public void SetControlEnabled1(System.Windows.Forms.Control.ControlCollection originalControls, bool EnabledFlag)
        {
       
            for (int i = 0; i <= originalControls.Count - 1; i++)
            {
                if (originalControls[i].Controls.Count > 0)
                {

                    SetControlEnabled1(originalControls[i].Controls, EnabledFlag);
                }


    

                try
                {
                    if (originalControls[i] is DataGridView)
                    {


                        DataGridView DataGridView = (DataGridView)originalControls[i];


                        DataGridView.ReadOnly = true;
                    }

                    if (originalControls[i] is Button)
                    {


                        Button aTextBox = (Button)originalControls[i];

        
                        aTextBox.Enabled = EnabledFlag;
                     
                    }

                }
                catch
                {
                }

            }

        }

        //�]�w Enter Leave �ƥ�B�z
        public void SetControlEvent(System.Windows.Forms.Control.ControlCollection originalControls)
        {
            // Control aControl;

            for (int i = 0; i <= originalControls.Count - 1; i++)
            {
                if (originalControls[i].Controls.Count > 0)
                {

                    SetControlEvent(originalControls[i].Controls);
                }

                if (originalControls[i] is TextBox)
                {


                    TextBox aTextBox = (TextBox)originalControls[i];

                    aTextBox.Enter += new EventHandler(TextBox_Enter);
                    aTextBox.Leave += new EventHandler(TextBox_Leave);
                }

                //�[�J 20071022 ComboBox
                if (originalControls[i] is ComboBox)
                {


                    ComboBox aTextBox = (ComboBox)originalControls[i];

                    aTextBox.Enter += new EventHandler(ComboBox_Enter);
                    aTextBox.Leave += new EventHandler(ComboBox_Leave);
                }
            }
        }

       


        //�i�J�J�I��,�ק��C��
        private void TextBox_Enter(object sender, EventArgs e)
        {
            ((TextBox)sender).BackColor = Color.Yellow;
        }

        //���}�J�I��,�ק��C��
        private void TextBox_Leave(object sender, EventArgs e)
        {
           // ((TextBox)sender).BackColor = Color.LightGray;
            ((TextBox)sender).BackColor = Color.White;
        }

        //�i�J�J�I��,�ק��C��
        private void ComboBox_Enter(object sender, EventArgs e)
        {
            ((ComboBox)sender).BackColor = Color.Yellow;
        }

        //���}�J�I��,�ק��C��
        private void ComboBox_Leave(object sender, EventArgs e)
        {
            // ((TextBox)sender).BackColor = Color.LightGray;
            ((ComboBox)sender).BackColor = Color.White;
        }

        //�M�� TextBox
        private void SetControlText(System.Windows.Forms.Control.ControlCollection originalControls)
        {
            Control aControl;

            for (int i = 0; i <= originalControls.Count - 1; i++)
            {
                if (originalControls[i].Controls.Count > 0)
                {

                    SetControlText(originalControls[i].Controls);
                }

                if (originalControls[i] is TextBox)
                {

                    // ((TextBox)originalControls[i]).BackColor = aColor;

                    TextBox aTextBox = (TextBox)originalControls[i];

                    aTextBox.Text = "";
                }
            }
        }

        private void bnQuery_Click(object sender, EventArgs e)
        {


            this.MyBS.AddNew();
            this.MyTableStatus = "9";
            SL_Status.Text = "�d��";
            this.SetControls();
        
            this.SetButtons();
            SetControlEnabled(Controls, true);
            SetControlEnabled1(Controls, false);
            Copy2.Enabled = false;
            this.query();

            SearchTextBox.Visible = false;
            bnSearch.Visible = false;
        }


        private void SetControlWhere(System.Windows.Forms.Control.ControlCollection originalControls)
        {
            Control aControl;

            for (int i = 0; i <= originalControls.Count - 1; i++)
            {
                if (originalControls[i].Controls.Count > 0)
                {

                    SetControlWhere(originalControls[i].Controls);
                }

                if (originalControls[i] is TextBox)
                {

                    // ((TextBox)originalControls[i]).BackColor = aColor;

                    TextBox aTextBox = (TextBox)originalControls[i];

                    string FieldName = GetRight(aTextBox.Name, "TextBox");

                    if (!string.IsNullOrEmpty(FieldName) && !string.IsNullOrEmpty(aTextBox.Text) && aTextBox.Text.Trim() != "")
                    {
                        //�ƭ��������Ҽ{

                        QW.Add(FieldName, MyTableName, TypeOfValues.StringType, WhereConditions.LikeAs, aTextBox, null);

                        //  QW.Add(
                    }
                } 
                else if  (originalControls[i] is ComboBox)
                {

                    // ((TextBox)originalControls[i]).BackColor = aColor;

                    ComboBox aComboBox = (ComboBox)originalControls[i];

                    string FieldName = GetRight(aComboBox.Name, "ComboBox");

                    WhereControl aWhereControl = new WhereControl(aComboBox, aComboBox, true, true);
                    aWhereControl.TableName = MyTableName;
                    aWhereControl.FieldName = FieldName;
                    aWhereControl.WhereCondition = WhereConditions.LikeAs;

                    //aWhereControl.Control1 = aComboBox.Name;
                    aWhereControl.SelectedValueIsUsed =true;

                    if (!string.IsNullOrEmpty(FieldName) && !string.IsNullOrEmpty(aComboBox.Text) && aComboBox.Text.Trim() != "")
                    {
                        //�ƭ��������Ҽ{

                      //  QW.Add(FieldName, MyTableName, TypeOfValues.StringType, WhereConditions.LikeAs, aComboBox, null);

                        QW.Add(aWhereControl);
                    }
                }
            }
        }

        private string GetRight(string s, string value)
        {
            int iIndex = s.IndexOf(value);

            if (iIndex >= 0)
            {

                return s.Substring(0, iIndex);

            }
            else
            {

                return string.Empty;
            }


        }

        private void ShowRecordCount(int RecordNo,int RecordCount)
        {
            SL_RecordCount.Text = string.Format("{0}/{1}", RecordNo, RecordCount);
        }

        private void bnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        //�v��
        // USERID, MENUID, ENABLED, APPEND, EDIT, DEL, REPORT
        private DataTable GetUSERMENUS(string USERID, string MENUID)
        {
            SqlConnection connection = new SqlConnection(strCn);

            string sql = "SELECT DISTINCT APPEND,EDIT,DEL,REPORT,SYSFLAG FROM USERMENUS " +
                "WHERE USERID IN (" + USERID + "  )  AND MENUID=@MENUID ORDER BY EDIT DESC,REPORT DESC";

            SqlCommand command = new SqlCommand(sql, connection);
            command.CommandType = CommandType.Text;
            command.Parameters.Add(new SqlParameter("@USERID", USERID));
            command.Parameters.Add(new SqlParameter("@MENUID", MENUID));
            SqlDataAdapter da = new SqlDataAdapter(command);

            System.Data.DataSet ds = new System.Data.DataSet();
            try
            {
                connection.Open();
                da.Fill(ds, "MENUS");
            }
            finally
            {
                connection.Close();
            }
            return ds.Tables["MENUS"];
        }

       
      


        private void Copy2_Click(object sender, EventArgs e)
        {
            Copy2.Enabled = false;

            SearchTextBox.Visible = false;
            bnSearch.Visible = false;
            //�ƻs
            this.AfterCopy();
            if (MasterTable == null)
            {
                MessageBox.Show("MasterTable �����]�w!!");
                return;
            }

            string Key = string.Empty;
            //if (DetailTables != null)
            //{
                //�}�� 
                
             
                    //string NumberName = "WH" + DateTime.Now.ToString("yyyyMMdd");
                    //string AutoNum = util.GetAutoNumber(MyConnection, NumberName);
                    ////this.shippingCodeTextBox.Text = NumberName + AutoNum + "X";
                Key = kyes;
                    //return;

                    //CloneTables = 
                
           // }
            DataTable[] dtArray = null;
                //�����ɭn���ƻs--------------------------------------------------------------------------
                // Master �� 0 ���ƻs
            if (DetailTables != null)
            {
                dtArray = new DataTable[DetailTables.Length];

                for (int i = 0; i <= DetailTables.Length - 1; i++)
                {
                    dtArray[i] = DetailTables[i].Copy();
                    DetailTables[i].Clear();
                }
            }
            //--------------------------------------------------------------------------


            //���ƻs�ܼȦs��
            DataTable dtTemp = MasterTable.Copy();

            //�s��s�W�|�y�� MasterTable �����ƫ���W�[
            //���I:�ƻs��������h�|��ܪŸ��
            //�t�@�Ӥ覡�ק�
            // CurrentRow[column.ColumnName] = dtTemp.Rows[0][column.ColumnName];
            //�令 CurrentRow[column.ColumnName] = dtTemp.Rows[dtTemp.Rows.Count-1][column.ColumnName];
            MasterTable.Clear();

            //MessageBox.Show(dtTemp.Rows.Count.ToString());
            //MessageBox.Show(Convert.ToString(dtTemp.Rows[0]["MENUID"]));
            //Ĳ�o�s�W
            bnAddNew.PerformClick();

      
            try
            {

                //�ϥ� DataRowView �t�X ResetCurrentItem �~����k���`�B�@
                DataRowView CurrentRow = (DataRowView)MyBS.Current;
                // CurrentRow["MENUID"] = Convert.ToString(dtTemp.Rows[0]["MENUID"]);
                foreach (DataColumn column in MasterTable.Columns)
                {

                    if (MyIDFieldName == column.ColumnName)
                    {
                        CurrentRow[column.ColumnName] = Key;
                    }
                    else
                    {
                        string S1 = column.ColumnName.ToString();
                        if (S1 != "NotifyMemo" && S1 != "MEMO1" && S1 != "MEMO2")
                        {
                            CurrentRow[column.ColumnName] = dtTemp.Rows[0][column.ColumnName];
                        }
                      
 
                    }
                }

                MyBS.ResetCurrentItem();
            }

            catch (Exception ex)
            { 

            }
            //�B�z������
            // && 
            // & �C�� �B�⦡�ҧ@
            if (DetailTables != null)
            {



                //���]�w������
                //�����ɭn���ƻs
                // if (DetailTables.Count > 0)
                if (DetailTables.Length > 0)
                {
                    int i = 0; //�]�� 0 �O Master

                    // �]����ȩ|���M�w
                    //�ҥH��ϥΪ̧���Ȯ�,�h�ݭn....
                    //�� Show �@�ӹ�ܲ�..��J���

                    //  foreach (DataTable table in DetailTables)
                    foreach (DataTable table in dtArray)
                    {

                        foreach (DataRow row in table.Rows)
                        {

                            DataRowView DetaiRow = (DataRowView)DetailBindingSources[i].AddNew();

                            // DataRowView DetaiRow = (DataRowView)DetailBindingSources[i].Current;

                            foreach (DataColumn column in table.Columns)
                            {

                                //��Ȥ��ݭn�ۤv�g�J
                                if (MyIDFieldName == column.ColumnName)
                                {
                                    DetaiRow[column.ColumnName] = Key;
                                }
                                else
                                {
                                    DetaiRow[column.ColumnName] = row[column.ColumnName];
                                }
                            }


                        }

                        i++;
                    }
                }
            }

                this.AfterCopy2();
        }

        private void SAVEButton_Click(object sender, EventArgs e)
        {
            SSTOPID = "0";
            STOP();
            if (SSTOPID == "1")
            {
                return;
            }
            SAVE();
        }

  

  
     
    }
}
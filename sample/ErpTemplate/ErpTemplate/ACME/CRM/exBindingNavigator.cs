using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;

namespace ACME
{

    

    public interface ITableUpdate
    {
        void Fill();
        void Update();
    }
    
    
    public partial class exBindingNavigator : UserControl, IContainerControl
    {
        public exBindingNavigator()
        {
            InitializeComponent();
            this.Dock = DockStyle.Top;
         //   BindingNavigatorSaveItem.Click += new EventHandler(BindingNavigatorSaveItem_Click);
         //   BindingNavigatorCancelItem.Click += new EventHandler(BindingNavigatorCancelItem_Click);
        }

        ITableUpdate tbl ;

        Form  _Form ; 
        Boolean _IsDataDirty = false;
        Boolean _AutoFillFlag = true;
        DataTable _DataTable;
        String _DisplayMember;
        BindingSource _BindingSource;
        Boolean _AutoSaveFlag = false;
        BindingSource _ParentBindingSource;

        private string FormStatus;


        public Boolean AutoSaveFlag
        {
            get { return _AutoSaveFlag; }
            set { _AutoSaveFlag = value; }
        }

        public Boolean IsDataDirty
        {
            get { return _IsDataDirty; }
            set 
            { 
                _IsDataDirty = value;
                if (BindingNavigatorSaveItem.Enabled != _IsDataDirty)
                    BindingNavigatorSaveItem.Enabled = _IsDataDirty;

                if (BindingNavigatorCancelItem.Enabled != _IsDataDirty)
                    BindingNavigatorCancelItem.Enabled = _IsDataDirty;

                //if (bindingNavigatorAddNewItem.Enabled != _IsDataDirty)
                //    bindingNavigatorAddNewItem.Enabled = _IsDataDirty;

                if (_IsDataDirty == false)
                {
                    bindingNavigatorAddNewItem.Enabled = true;
                    bindingNavigatorDeleteItem.Enabled = true;
                }
                else
                {
                    bindingNavigatorAddNewItem.Enabled = false;
                    bindingNavigatorDeleteItem.Enabled = false;
                }

            }
        }

        public Boolean AutoFillFlag
        {
            get { return _AutoFillFlag; }
            set { _AutoFillFlag = value; }
        }

        public DataTable DataTable
        {
            get { return _DataTable; }
            set { _DataTable = value; }
        }

        public BindingSource BindingSource
        {
            get { return _BindingSource; }
            set 
            { 
                GenericBindingNavigator.BindingSource = value;
                _BindingSource = value;
                if (_BindingSource != null)
                {
                    ((System.ComponentModel.ISupportInitialize)(_BindingSource)).BeginInit();
                    //subscribe to the events in case not yet set
                    _BindingSource.DataSourceChanged
                        += new EventHandler(bs_DataSourceChanged);
                    _BindingSource.DataMemberChanged
                        += new EventHandler(bs_DataSourceChanged);
                    _BindingSource.PositionChanged
                        += new EventHandler(bs_PositionChanged);

                    //不知的事件
                    //利用此事件的判斷是否在編輯狀態
                    _BindingSource.BindingComplete
                        += new BindingCompleteEventHandler(BindingComplete);
                    _BindingSource.ListChanged
                        += new ListChangedEventHandler(bs_ListChanged);
                    bs_DataSourceChanged(new object(), new EventArgs());
                    ((System.ComponentModel.ISupportInitialize)(_BindingSource)).EndInit();

                }
            }
        }            
        public String DisplayMember
        {
            get { return _DisplayMember; }
            set { _DisplayMember = value; }
        }


        public BindingSource ParentBindingSource
        {
            get { return _ParentBindingSource; }
            set
            {
                _ParentBindingSource = value;
                if (_ParentBindingSource != null)
                {
                    //subscribe to the position changed event to rebuild the lookup list
                    _ParentBindingSource.PositionChanged += new EventHandler(parentBS_PositionChanged);
                }
            }
        }
 //properties


        private void BindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            Validate();

            MyEventArgs a = new MyEventArgs();

            a.CheckOk = true;
            if (BeforePost != null)
            {
                
                BeforePost(sender, a)  ;
            }

            if (a.CheckOk != true)
            {
                return;
            }

            try
            {
                
                GenericBindingNavigator.BindingSource.EndEdit();
                tbl = (_DataTable as ITableUpdate);
                tbl.Update();

                _DataTable.AcceptChanges();

                IsDataDirty = false;

               // MessageBox.Show(FormStatus);
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "操作錯誤", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                return;
            }



           
            if (AfterSave != null)
            {

                MyFormStatusEventArgs arg = new MyFormStatusEventArgs();

                arg.MyFormStatus = FormStatus;

                AfterSave(sender, arg);
            }

            FormStatus = "R";


          
        }

        private void BindingNavigatorCancelItem_Click(object sender, EventArgs e)
        {

           

            Validate();

            //取消前詢問
            if (_IsDataDirty & _DataTable != null)
            {
                String msg = "Do you want to save edits to the previous record?";
                if (_AutoSaveFlag | MessageBox.Show(msg, "Confirm Save", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    BindingNavigatorSaveItem_Click(new Object(), new EventArgs());

                    return;
                }
                else
                {
                    IsDataDirty = false;
                    GenericBindingNavigator.BindingSource.CancelEdit();
                    _DataTable.RejectChanges();

                    FormStatus = "R";
                    //  MessageBox.Show("All unsaved edits have been rolled back.");
                    //IsDataDirty = false;
                }
            }

            

            
        }

        private void Form_Load(Object sender, EventArgs e)
        {
            bs_DataSourceChanged(new object(), new EventArgs());
            if (_AutoFillFlag)
            {
                tbl = (_DataTable as ITableUpdate); //Cast as interface
                tbl.Fill();
            }
        }

        //利用這個事件來判斷是 編輯
        private void BindingComplete(Object sender, BindingCompleteEventArgs e)
        {
            if (e.BindingCompleteContext == BindingCompleteContext.DataSourceUpdate)
                if (e.BindingCompleteState == BindingCompleteState.Success &
                        !e.Binding.Control.BindingContext.IsReadOnly)
                {
                    IsDataDirty = true;

                    if (FormStatus == "I")
                    {
                        return;
                    }

                    FormStatus = "E";
                }
        }
        private void bs_DataSourceChanged(object sender, EventArgs e)
        {
            DataTable PrevTable = _DataTable;
            _DataTable = GetTableFromBindingSource(GenericBindingNavigator.BindingSource);
            if (_DataTable != null & PrevTable!=_DataTable)
            {
                if (_DisplayMember != null)
                {

                    //find the first text column
                    foreach (DataColumn col in _DataTable.Columns)
                    {
                        if (col.DataType == typeof(string))
                        {
                            _DisplayMember = col.ColumnName;
                            break;
                        }
                    }
                    //if child BS, get reference to parent BS
                    BindingSource testBS = (GenericBindingNavigator.BindingSource.DataSource as BindingSource);
                    if (testBS != null)
                    {
                        ParentBindingSource = testBS; //call the get to capture event
                    }
                }
            }
        }
        //handles the Position Changed for the binding source
        //資料移動時不處理...太難了
        //位置改變了...如何....
        //使用者移動 dataGridView
        private void bs_PositionChanged(object sender, EventArgs e)
        {
            //if (_IsDataDirty & _DataTable !=null) 
            //{
            //    String msg = "Do you want to save edits to the previous record?";
            //    if (_AutoSaveFlag | MessageBox.Show(msg, "Confirm Save", MessageBoxButtons.YesNo) == DialogResult.Yes )
            //    {
            //        BindingNavigatorSaveItem_Click(new Object(), new EventArgs());
                   
            //    }
            //    else
            //    {
            //        _DataTable.RejectChanges();
            //        //  MessageBox.Show("All unsaved edits have been rolled back.");
            //        IsDataDirty = false;
            //    }
            //}


            if (_IsDataDirty & _DataTable != null)
            {
                
                 //   _DataTable.RejectChanges();
                    //  MessageBox.Show("All unsaved edits have been rolled back.");
                    IsDataDirty = false;
                    
                
            }

            
           // syncLookupCombobox();
        }

        private void parentBS_PositionChanged(object sender, EventArgs e)
        {
            BuildLookupList();
        }
        private void bs_ListChanged(object sender, ListChangedEventArgs e)
        {
            if(e.ListChangedType==ListChangedType.Reset)
                BuildLookupList();
        }

        private void exBindingNavigator_Load(object sender, EventArgs e)
        {
            //get the reference to the hosting form
            //if (_Form == null)
            //{
            //    Object frm = (this as ContainerControl).ParentForm;
            //    while (frm != null && (frm as Form) == null)
            //    {
            //        if ((frm as ContainerControl) != null)
            //            frm = (frm as ContainerControl).ParentForm;
            //        else if ((frm as Control) != null)
            //            frm = (frm as Control).Parent;
            //    }
            //    if (frm != null)
            //    {
            //        _Form = (frm as Form);
            //        _Form.Load += new EventHandler(Form_Load);
            //    }
            //}

            if (!DesignMode)
            {

                bs_DataSourceChanged(new object(), new EventArgs());
                if (_AutoFillFlag)
                {
                    try
                    {
                        tbl = (_DataTable as ITableUpdate); //Cast as interface

                        tbl.Fill();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }


        private System.Collections.Hashtable ht = new System.Collections.Hashtable();

        public void BuildLookupList()
        {
            //if (_DisplayMember != null)
            //{
            //    //fill both lookup box and hash table with values
            //    toolStripComboBox1.Items.Clear();
            //    ht.Clear();
            //    if (BindingSource.List.Count > 0)
            //    {
            //        //get the primary key as value member (assumes single field key)
            //        string _valueMember = _DataTable.PrimaryKey[0].ColumnName;
            //        //temp change the sort of the binding source
            //        string tempSort = _BindingSource.Sort;
            //        BindingSource.Sort = _DisplayMember + " ASC";
            //        int cnt = 0;
            //        //step through the records in the binding source as filtered
            //        foreach (DataRowView drv in _BindingSource)
            //        {
            //            if (drv[_DisplayMember] != null)
            //            {
            //                toolStripComboBox1.Items.Add(drv[_DisplayMember]);
            //                try
            //                {
            //                    ht.Add(drv[_DisplayMember], drv[_valueMember]);
            //                }
            //                catch { } //ignore dups
            //            }
            //            cnt += 1;
            //        }
            //        //restore sort field
            //        BindingSource.Sort = tempSort;
            //        syncLookupCombobox();
            //    }
            //}
        }

        /// <summary>
        /// Get the name of the table behind a Binding Source
        /// </summary>
        /// <param name="bs">The binding source</param>
        /// <returns>DataTable instance</returns>
        public DataTable GetTableFromBindingSource(BindingSource bs)
        {
            if (bs==null || bs.DataSource == null | bs.DataMember == "")
                return null;
            System.Data.DataSet _dataSet;
            System.Data.DataTable _dataTable;
            object obj = bs.DataSource;
            //if datasource is another binding source, loop until the parent dataset is found.
            while ((obj as BindingSource) != null)
                obj = (obj as BindingSource).DataSource;
            //make sure obj is now a dataset
            if ((obj as DataSet) == null)
                return null;
            _dataSet = (DataSet)obj;
            //is the DataMember a table or a relation
            _dataTable = (_dataSet.Tables[bs.DataMember] as DataTable);
            if (_dataTable == null)
                //it must be a relation instead of a table
                _dataTable = _dataSet.Relations[bs.DataMember].ChildTable as DataTable;

            //if (_dataTable != null)
            //{
            //    //make sure the table implements the interface
            //    if ((_dataTable as ITableUpdate) == null)
            //    {
            //        throw new ApplicationException("資料表 " + _dataTable.TableName +
            //            " does not implement ITableUpdate interface");
            //    }
            //    return _dataTable;
            //}
            //else
            //    return null;

            if (_dataTable != null)
            {
               
                return _dataTable;
            }
            else
                return null;
        }

         //handler for delete button press
        private void bindingNavigatorDeleteItem_Click(object sender, EventArgs e)
        {
            FormStatus = "D";

            MyEventArgs a = new MyEventArgs();
            a.CheckOk = true;

            if (BeforeDelete != null)
            {
              
                BeforeDelete(sender, a);
            }

            if (a.CheckOk != true)
            {
                return;
            }



            //string msg = "Are you sure you want to delete the current record? " + Environment.NewLine;
            string msg = "Are you sure you want to delete the current record? ";
           //  + "ID=" +(this.BindingSource[BindingSource.Position] as DataRowView).Row[0].ToString();
            if (_AutoSaveFlag || MessageBox.Show(msg, "Confirm Delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                //Delete the current record -
                this.BindingSource.RemoveCurrent();
                (_DataTable as ITableUpdate).Update();
                _DataTable.AcceptChanges();
              //  BindingNavigatorSaveItem_Click(sender, e);
            }
        }

        //可以拿來用尋找
        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //if (toolStripComboBox1.Focused && toolStripComboBox1.Text != "")
            //{
            //    //get the primary key field
            //    string _valueMember = _DataTable.PrimaryKey[0].ColumnName;
            //    //set the position by finding the primary key in the binding source
            //    _BindingSource.Position = BindingSource.Find(_valueMember, ht[toolStripComboBox1.Text]);
            //}
        }

        //make sure the same record shows in the combo box as is in the record
        private void syncLookupCombobox()
        {
            //if (toolStripComboBox1.Items.Count > 0 & _BindingSource.Position >= 0)
            //{
            //    //get the display string for the current record in the binding source
            //    string lookup = (_BindingSource.Current as DataRowView).Row[_DisplayMember].ToString();
            //    if (lookup.Length > 0)
            //        toolStripComboBox1.SelectedIndex = toolStripComboBox1.FindStringExact(lookup);
            //}
        }

        private void bindingNavigatorAddNewItem_Click(object sender, EventArgs e)
        {
            FormStatus = "I";
            
            Validate();
            GenericBindingNavigator.BindingSource.AddNew();

            if (AfterNew != null)
                AfterNew(sender, e);

            IsDataDirty = true;

        }



        public delegate void BeforeDeleteEventHandler(object sender, MyEventArgs args);
        public event BeforeDeleteEventHandler BeforeDelete;

        public delegate void BeforePostEventHandler(object sender, MyEventArgs args);
        public event BeforePostEventHandler BeforePost;

        public delegate void AfterNewEventHandler(object sender, EventArgs e);
        public event AfterNewEventHandler AfterNew;

        //AfterSave
        public delegate void AfterSaveEventHandler(object sender, MyFormStatusEventArgs args);
        public event AfterSaveEventHandler AfterSave;

        

        //private void OnTextClicked(object sender, EventArgs args)
        //{
        //    //Call the delegate
        //    if (TextClicked != null)
        //        TextClicked(sender, args);
        //}
 
    }

    public class MyEventArgs : EventArgs
    {
        public Boolean _CheckOk;

        public Boolean CheckOk
        {
            get { return _CheckOk; }
            set { _CheckOk = value; }
        }

        public MyEventArgs()
        {

        }
    }

    public class MyFormStatusEventArgs : EventArgs
    {
        public string _MyFormStatus;

        public string MyFormStatus
        {
            get { return _MyFormStatus; }
            set { _MyFormStatus = value; }
        }

        public MyFormStatusEventArgs()
        {

        }
    }
}



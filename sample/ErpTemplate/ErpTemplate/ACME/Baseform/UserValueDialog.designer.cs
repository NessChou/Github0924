namespace ACME
{
	partial class UserValueDialog
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UserValueDialog));
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.tbExpression = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnSave = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.aCME_UserValueDataGridView = new System.Windows.Forms.DataGridView();
            this.FormID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ObjID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.KeyValue = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.aCME_UserValueBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.userValue = new ACME.Baseform.UserValue();
            this.aCME_UserValueBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorDeleteItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.aCME_UserValueTableAdapter = new ACME.Baseform.UserValueTableAdapters.ACME_UserValueTableAdapter();
            this.groupBox1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_UserValueDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_UserValueBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.userValue)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_UserValueBindingNavigator)).BeginInit();
            this.aCME_UserValueBindingNavigator.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.Name = "btnOK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.tbExpression);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // tbExpression
            // 
            resources.ApplyResources(this.tbExpression, "tbExpression");
            this.tbExpression.Name = "tbExpression";
            this.tbExpression.TextChanged += new System.EventHandler(this.tbExpression_TextChanged);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnSave);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.btnOK);
            resources.ApplyResources(this.panel1, "panel1");
            this.panel1.Name = "panel1";
            // 
            // btnSave
            // 
            resources.ApplyResources(this.btnSave, "btnSave");
            this.btnSave.Name = "btnSave";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.groupBox3);
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.aCME_UserValueDataGridView);
            resources.ApplyResources(this.groupBox3, "groupBox3");
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.TabStop = false;
            // 
            // aCME_UserValueDataGridView
            // 
            this.aCME_UserValueDataGridView.AutoGenerateColumns = false;
            this.aCME_UserValueDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.FormID,
            this.ObjID,
            this.KeyValue});
            this.aCME_UserValueDataGridView.DataSource = this.aCME_UserValueBindingSource;
            resources.ApplyResources(this.aCME_UserValueDataGridView, "aCME_UserValueDataGridView");
            this.aCME_UserValueDataGridView.Name = "aCME_UserValueDataGridView";
            this.aCME_UserValueDataGridView.RowTemplate.Height = 24;
            this.aCME_UserValueDataGridView.DoubleClick += new System.EventHandler(this.aCME_UserValueDataGridView_DoubleClick);
            this.aCME_UserValueDataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.aCME_UserValueDataGridView_DefaultValuesNeeded);
            // 
            // FormID
            // 
            this.FormID.DataPropertyName = "FormID";
            resources.ApplyResources(this.FormID, "FormID");
            this.FormID.Name = "FormID";
            // 
            // ObjID
            // 
            this.ObjID.DataPropertyName = "ObjID";
            resources.ApplyResources(this.ObjID, "ObjID");
            this.ObjID.Name = "ObjID";
            // 
            // KeyValue
            // 
            this.KeyValue.DataPropertyName = "KeyValue";
            resources.ApplyResources(this.KeyValue, "KeyValue");
            this.KeyValue.Name = "KeyValue";
            // 
            // aCME_UserValueBindingSource
            // 
            this.aCME_UserValueBindingSource.DataMember = "ACME_UserValue";
            this.aCME_UserValueBindingSource.DataSource = this.userValue;
            // 
            // userValue
            // 
            this.userValue.DataSetName = "UserValue";
            this.userValue.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // aCME_UserValueBindingNavigator
            // 
            this.aCME_UserValueBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.aCME_UserValueBindingNavigator.BindingSource = this.aCME_UserValueBindingSource;
            this.aCME_UserValueBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.aCME_UserValueBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.aCME_UserValueBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bindingNavigatorMoveFirstItem,
            this.bindingNavigatorMovePreviousItem,
            this.bindingNavigatorSeparator,
            this.bindingNavigatorPositionItem,
            this.bindingNavigatorCountItem,
            this.bindingNavigatorSeparator1,
            this.bindingNavigatorMoveNextItem,
            this.bindingNavigatorMoveLastItem,
            this.bindingNavigatorSeparator2,
            this.bindingNavigatorAddNewItem,
            this.bindingNavigatorDeleteItem});
            resources.ApplyResources(this.aCME_UserValueBindingNavigator, "aCME_UserValueBindingNavigator");
            this.aCME_UserValueBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.aCME_UserValueBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.aCME_UserValueBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.aCME_UserValueBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.aCME_UserValueBindingNavigator.Name = "aCME_UserValueBindingNavigator";
            this.aCME_UserValueBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            // 
            // bindingNavigatorAddNewItem
            // 
            this.bindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.bindingNavigatorAddNewItem, "bindingNavigatorAddNewItem");
            this.bindingNavigatorAddNewItem.Name = "bindingNavigatorAddNewItem";
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            resources.ApplyResources(this.bindingNavigatorCountItem, "bindingNavigatorCountItem");
            // 
            // bindingNavigatorDeleteItem
            // 
            this.bindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.bindingNavigatorDeleteItem, "bindingNavigatorDeleteItem");
            this.bindingNavigatorDeleteItem.Name = "bindingNavigatorDeleteItem";
            // 
            // bindingNavigatorMoveFirstItem
            // 
            this.bindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.bindingNavigatorMoveFirstItem, "bindingNavigatorMoveFirstItem");
            this.bindingNavigatorMoveFirstItem.Name = "bindingNavigatorMoveFirstItem";
            // 
            // bindingNavigatorMovePreviousItem
            // 
            this.bindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.bindingNavigatorMovePreviousItem, "bindingNavigatorMovePreviousItem");
            this.bindingNavigatorMovePreviousItem.Name = "bindingNavigatorMovePreviousItem";
            // 
            // bindingNavigatorSeparator
            // 
            this.bindingNavigatorSeparator.Name = "bindingNavigatorSeparator";
            resources.ApplyResources(this.bindingNavigatorSeparator, "bindingNavigatorSeparator");
            // 
            // bindingNavigatorPositionItem
            // 
            resources.ApplyResources(this.bindingNavigatorPositionItem, "bindingNavigatorPositionItem");
            this.bindingNavigatorPositionItem.Name = "bindingNavigatorPositionItem";
            // 
            // bindingNavigatorSeparator1
            // 
            this.bindingNavigatorSeparator1.Name = "bindingNavigatorSeparator1";
            resources.ApplyResources(this.bindingNavigatorSeparator1, "bindingNavigatorSeparator1");
            // 
            // bindingNavigatorMoveNextItem
            // 
            this.bindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.bindingNavigatorMoveNextItem, "bindingNavigatorMoveNextItem");
            this.bindingNavigatorMoveNextItem.Name = "bindingNavigatorMoveNextItem";
            // 
            // bindingNavigatorMoveLastItem
            // 
            this.bindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            resources.ApplyResources(this.bindingNavigatorMoveLastItem, "bindingNavigatorMoveLastItem");
            this.bindingNavigatorMoveLastItem.Name = "bindingNavigatorMoveLastItem";
            // 
            // bindingNavigatorSeparator2
            // 
            this.bindingNavigatorSeparator2.Name = "bindingNavigatorSeparator2";
            resources.ApplyResources(this.bindingNavigatorSeparator2, "bindingNavigatorSeparator2");
            // 
            // aCME_UserValueTableAdapter
            // 
            this.aCME_UserValueTableAdapter.ClearBeforeFill = true;
            // 
            // UserValueDialog
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.aCME_UserValueBindingNavigator);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UserValueDialog";
            this.Load += new System.EventHandler(this.LookupDialog_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.panel1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.aCME_UserValueDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_UserValueBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.userValue)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aCME_UserValueBindingNavigator)).EndInit();
            this.aCME_UserValueBindingNavigator.ResumeLayout(false);
            this.aCME_UserValueBindingNavigator.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox tbExpression;
        private System.Windows.Forms.Panel panel1;
        public System.Windows.Forms.GroupBox groupBox2;
        public System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label1;
        private ACME.Baseform.UserValue userValue;
        private System.Windows.Forms.BindingSource aCME_UserValueBindingSource;
        private ACME.Baseform.UserValueTableAdapters.ACME_UserValueTableAdapter aCME_UserValueTableAdapter;
        private System.Windows.Forms.BindingNavigator aCME_UserValueBindingNavigator;
        private System.Windows.Forms.ToolStripButton bindingNavigatorAddNewItem;
        private System.Windows.Forms.ToolStripLabel bindingNavigatorCountItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorDeleteItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveFirstItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMovePreviousItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator;
        private System.Windows.Forms.ToolStripTextBox bindingNavigatorPositionItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator1;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveNextItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveLastItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.DataGridView aCME_UserValueDataGridView;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.DataGridViewTextBoxColumn FormID;
        private System.Windows.Forms.DataGridViewTextBoxColumn ObjID;
        private System.Windows.Forms.DataGridViewTextBoxColumn KeyValue;
        
	}
}
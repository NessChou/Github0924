namespace ACME
{
    partial class GB_FRIGHT
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GB_FRIGHT));
            this.gB_FRIGHTBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.gB_FRIGHTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.pOTATO = new ACME.ACMEDataSet.POTATO();
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
            this.gB_FRIGHTBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.gB_FRIGHTDataGridView = new System.Windows.Forms.DataGridView();
            this.gB_FRIGHTTableAdapter = new ACME.ACMEDataSet.POTATOTableAdapters.GB_FRIGHTTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.POTATOTableAdapters.TableAdapterManager();
            this.CELLNAME = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FRIGHTBindingNavigator)).BeginInit();
            this.gB_FRIGHTBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FRIGHTBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pOTATO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FRIGHTDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // gB_FRIGHTBindingNavigator
            // 
            this.gB_FRIGHTBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.gB_FRIGHTBindingNavigator.BindingSource = this.gB_FRIGHTBindingSource;
            this.gB_FRIGHTBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.gB_FRIGHTBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.gB_FRIGHTBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.bindingNavigatorDeleteItem,
            this.gB_FRIGHTBindingNavigatorSaveItem});
            this.gB_FRIGHTBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.gB_FRIGHTBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.gB_FRIGHTBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.gB_FRIGHTBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.gB_FRIGHTBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.gB_FRIGHTBindingNavigator.Name = "gB_FRIGHTBindingNavigator";
            this.gB_FRIGHTBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.gB_FRIGHTBindingNavigator.Size = new System.Drawing.Size(952, 25);
            this.gB_FRIGHTBindingNavigator.TabIndex = 0;
            this.gB_FRIGHTBindingNavigator.Text = "bindingNavigator1";
            // 
            // bindingNavigatorAddNewItem
            // 
            this.bindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorAddNewItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorAddNewItem.Image")));
            this.bindingNavigatorAddNewItem.Name = "bindingNavigatorAddNewItem";
            this.bindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorAddNewItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorAddNewItem.Text = "加入新的";
            // 
            // gB_FRIGHTBindingSource
            // 
            this.gB_FRIGHTBindingSource.DataMember = "GB_FRIGHT";
            this.gB_FRIGHTBindingSource.DataSource = this.pOTATO;
            // 
            // pOTATO
            // 
            this.pOTATO.DataSetName = "POTATO";
            this.pOTATO.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(28, 22);
            this.bindingNavigatorCountItem.Text = "/{0}";
            this.bindingNavigatorCountItem.ToolTipText = "項目總數";
            // 
            // bindingNavigatorDeleteItem
            // 
            this.bindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorDeleteItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorDeleteItem.Image")));
            this.bindingNavigatorDeleteItem.Name = "bindingNavigatorDeleteItem";
            this.bindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorDeleteItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorDeleteItem.Text = "刪除";
            // 
            // bindingNavigatorMoveFirstItem
            // 
            this.bindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveFirstItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveFirstItem.Image")));
            this.bindingNavigatorMoveFirstItem.Name = "bindingNavigatorMoveFirstItem";
            this.bindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveFirstItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveFirstItem.Text = "移到最前面";
            // 
            // bindingNavigatorMovePreviousItem
            // 
            this.bindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMovePreviousItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMovePreviousItem.Image")));
            this.bindingNavigatorMovePreviousItem.Name = "bindingNavigatorMovePreviousItem";
            this.bindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMovePreviousItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMovePreviousItem.Text = "移到上一個";
            // 
            // bindingNavigatorSeparator
            // 
            this.bindingNavigatorSeparator.Name = "bindingNavigatorSeparator";
            this.bindingNavigatorSeparator.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorPositionItem
            // 
            this.bindingNavigatorPositionItem.AccessibleName = "位置";
            this.bindingNavigatorPositionItem.AutoSize = false;
            this.bindingNavigatorPositionItem.Name = "bindingNavigatorPositionItem";
            this.bindingNavigatorPositionItem.Size = new System.Drawing.Size(50, 23);
            this.bindingNavigatorPositionItem.Text = "0";
            this.bindingNavigatorPositionItem.ToolTipText = "目前的位置";
            // 
            // bindingNavigatorSeparator1
            // 
            this.bindingNavigatorSeparator1.Name = "bindingNavigatorSeparator1";
            this.bindingNavigatorSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // bindingNavigatorMoveNextItem
            // 
            this.bindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveNextItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveNextItem.Image")));
            this.bindingNavigatorMoveNextItem.Name = "bindingNavigatorMoveNextItem";
            this.bindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveNextItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveNextItem.Text = "移到下一個";
            // 
            // bindingNavigatorMoveLastItem
            // 
            this.bindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveLastItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveLastItem.Image")));
            this.bindingNavigatorMoveLastItem.Name = "bindingNavigatorMoveLastItem";
            this.bindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveLastItem.Size = new System.Drawing.Size(23, 22);
            this.bindingNavigatorMoveLastItem.Text = "移到最後面";
            // 
            // bindingNavigatorSeparator2
            // 
            this.bindingNavigatorSeparator2.Name = "bindingNavigatorSeparator2";
            this.bindingNavigatorSeparator2.Size = new System.Drawing.Size(6, 25);
            // 
            // gB_FRIGHTBindingNavigatorSaveItem
            // 
            this.gB_FRIGHTBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.gB_FRIGHTBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("gB_FRIGHTBindingNavigatorSaveItem.Image")));
            this.gB_FRIGHTBindingNavigatorSaveItem.Name = "gB_FRIGHTBindingNavigatorSaveItem";
            this.gB_FRIGHTBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.gB_FRIGHTBindingNavigatorSaveItem.Text = "儲存資料";
            this.gB_FRIGHTBindingNavigatorSaveItem.Click += new System.EventHandler(this.gB_FRIGHTBindingNavigatorSaveItem_Click);
            // 
            // gB_FRIGHTDataGridView
            // 
            this.gB_FRIGHTDataGridView.AutoGenerateColumns = false;
            this.gB_FRIGHTDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.gB_FRIGHTDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gB_FRIGHTDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CELLNAME,
            this.dataGridViewTextBoxColumn2});
            this.gB_FRIGHTDataGridView.DataSource = this.gB_FRIGHTBindingSource;
            this.gB_FRIGHTDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gB_FRIGHTDataGridView.Location = new System.Drawing.Point(0, 25);
            this.gB_FRIGHTDataGridView.Name = "gB_FRIGHTDataGridView";
            this.gB_FRIGHTDataGridView.RowTemplate.Height = 24;
            this.gB_FRIGHTDataGridView.Size = new System.Drawing.Size(952, 552);
            this.gB_FRIGHTDataGridView.TabIndex = 2;
            this.gB_FRIGHTDataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.gB_FRIGHTDataGridView_DataError);
            // 
            // gB_FRIGHTTableAdapter
            // 
            this.gB_FRIGHTTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.GB_BOMDTableAdapter = null;
            this.tableAdapterManager.GB_BOMMTableAdapter = null;
            this.tableAdapterManager.GB_CS1TableAdapter = null;
            this.tableAdapterManager.GB_CS2TableAdapter = null;
            this.tableAdapterManager.GB_CSDTableAdapter = null;
            this.tableAdapterManager.GB_CSTableAdapter = null;
            this.tableAdapterManager.GB_DATELOCKTableAdapter = null;
            this.tableAdapterManager.GB_DEADLINETableAdapter = null;
            this.tableAdapterManager.GB_FISHTableAdapter = null;
            this.tableAdapterManager.GB_FOC2TableAdapter = null;
            this.tableAdapterManager.GB_FOC3TableAdapter = null;
            this.tableAdapterManager.GB_FOCTableAdapter = null;
            this.tableAdapterManager.GB_FPRODUCTTableAdapter = null;
            this.tableAdapterManager.GB_FRIEND1TableAdapter = null;
            this.tableAdapterManager.GB_FRIENDTableAdapter = null;
            this.tableAdapterManager.GB_FRIGHTTableAdapter = this.gB_FRIGHTTableAdapter;
            this.tableAdapterManager.GB_INVTRACKTableAdapter = null;
            this.tableAdapterManager.GB_OCRD2TableAdapter = null;
            this.tableAdapterManager.GB_OCRDTableAdapter = null;
            this.tableAdapterManager.GB_PICK2TableAdapter = null;
            this.tableAdapterManager.GB_PICKTableAdapter = null;
            this.tableAdapterManager.GB_POTATO1TableAdapter = null;
            this.tableAdapterManager.GB_POTATO21TableAdapter = null;
            this.tableAdapterManager.GB_POTATO2TableAdapter = null;
            this.tableAdapterManager.GB_POTATOTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.POTATOTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // CELLNAME
            // 
            this.CELLNAME.DataPropertyName = "CELLNAME";
            this.CELLNAME.HeaderText = "類型";
            this.CELLNAME.Items.AddRange(new object[] {
            "外站+Roots",
            "棉花田",
            "安永",
            "電話+傳真",
            "員購",
            "短效品",
            "官網",
            "批發",
            "大宗樣品",
            "其他銷貨",
            "預計進貨"});
            this.CELLNAME.Name = "CELLNAME";
            this.CELLNAME.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.CELLNAME.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.CELLNAME.Width = 54;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "USERID";
            this.dataGridViewTextBoxColumn2.HeaderText = "使用帳號";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 78;
            // 
            // GB_FRIGHT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(952, 577);
            this.Controls.Add(this.gB_FRIGHTDataGridView);
            this.Controls.Add(this.gB_FRIGHTBindingNavigator);
            this.Name = "GB_FRIGHT";
            this.Text = "權限設定";
            this.Load += new System.EventHandler(this.GB_FRIGHT_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gB_FRIGHTBindingNavigator)).EndInit();
            this.gB_FRIGHTBindingNavigator.ResumeLayout(false);
            this.gB_FRIGHTBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FRIGHTBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pOTATO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FRIGHTDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.POTATO pOTATO;
        private System.Windows.Forms.BindingSource gB_FRIGHTBindingSource;
        private ACMEDataSet.POTATOTableAdapters.GB_FRIGHTTableAdapter gB_FRIGHTTableAdapter;
        private ACMEDataSet.POTATOTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.BindingNavigator gB_FRIGHTBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton gB_FRIGHTBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView gB_FRIGHTDataGridView;
        private System.Windows.Forms.DataGridViewComboBoxColumn CELLNAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
    }
}
namespace ACME
{
    partial class GB_FPRODUCT
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GB_FPRODUCT));
            this.pOTATO = new ACME.ACMEDataSet.POTATO();
            this.gB_FPRODUCTBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gB_FPRODUCTTableAdapter = new ACME.ACMEDataSet.POTATOTableAdapters.GB_FPRODUCTTableAdapter();
            this.gB_FPRODUCTBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.gB_FPRODUCTBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.gB_FPRODUCTDataGridView = new System.Windows.Forms.DataGridView();
            this.ITEMCODE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ITEMNAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.UNIT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.tableAdapterManager = new ACME.ACMEDataSet.POTATOTableAdapters.TableAdapterManager();
            ((System.ComponentModel.ISupportInitialize)(this.pOTATO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FPRODUCTBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FPRODUCTBindingNavigator)).BeginInit();
            this.gB_FPRODUCTBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FPRODUCTDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // pOTATO
            // 
            this.pOTATO.DataSetName = "POTATO";
            this.pOTATO.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // gB_FPRODUCTBindingSource
            // 
            this.gB_FPRODUCTBindingSource.DataMember = "GB_FPRODUCT";
            this.gB_FPRODUCTBindingSource.DataSource = this.pOTATO;
            // 
            // gB_FPRODUCTTableAdapter
            // 
            this.gB_FPRODUCTTableAdapter.ClearBeforeFill = true;
            // 
            // gB_FPRODUCTBindingNavigator
            // 
            this.gB_FPRODUCTBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.gB_FPRODUCTBindingNavigator.BindingSource = this.gB_FPRODUCTBindingSource;
            this.gB_FPRODUCTBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.gB_FPRODUCTBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.gB_FPRODUCTBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.gB_FPRODUCTBindingNavigatorSaveItem});
            this.gB_FPRODUCTBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.gB_FPRODUCTBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.gB_FPRODUCTBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.gB_FPRODUCTBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.gB_FPRODUCTBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.gB_FPRODUCTBindingNavigator.Name = "gB_FPRODUCTBindingNavigator";
            this.gB_FPRODUCTBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.gB_FPRODUCTBindingNavigator.Size = new System.Drawing.Size(960, 25);
            this.gB_FPRODUCTBindingNavigator.TabIndex = 0;
            this.gB_FPRODUCTBindingNavigator.Text = "bindingNavigator1";
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
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(24, 22);
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
            // gB_FPRODUCTBindingNavigatorSaveItem
            // 
            this.gB_FPRODUCTBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.gB_FPRODUCTBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("gB_FPRODUCTBindingNavigatorSaveItem.Image")));
            this.gB_FPRODUCTBindingNavigatorSaveItem.Name = "gB_FPRODUCTBindingNavigatorSaveItem";
            this.gB_FPRODUCTBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.gB_FPRODUCTBindingNavigatorSaveItem.Text = "儲存資料";
            this.gB_FPRODUCTBindingNavigatorSaveItem.Click += new System.EventHandler(this.gB_FPRODUCTBindingNavigatorSaveItem_Click);
            // 
            // gB_FPRODUCTDataGridView
            // 
            this.gB_FPRODUCTDataGridView.AutoGenerateColumns = false;
            this.gB_FPRODUCTDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.gB_FPRODUCTDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gB_FPRODUCTDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ITEMCODE,
            this.ITEMNAME,
            this.UNIT,
            this.dataGridViewTextBoxColumn5});
            this.gB_FPRODUCTDataGridView.DataSource = this.gB_FPRODUCTBindingSource;
            this.gB_FPRODUCTDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gB_FPRODUCTDataGridView.Location = new System.Drawing.Point(0, 25);
            this.gB_FPRODUCTDataGridView.Name = "gB_FPRODUCTDataGridView";
            this.gB_FPRODUCTDataGridView.RowTemplate.Height = 24;
            this.gB_FPRODUCTDataGridView.Size = new System.Drawing.Size(960, 599);
            this.gB_FPRODUCTDataGridView.TabIndex = 1;
            this.gB_FPRODUCTDataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.gB_FPRODUCTDataGridView_CellValueChanged);
            // 
            // ITEMCODE
            // 
            this.ITEMCODE.DataPropertyName = "ITEMCODE";
            this.ITEMCODE.HeaderText = "料號";
            this.ITEMCODE.Name = "ITEMCODE";
            this.ITEMCODE.Width = 54;
            // 
            // ITEMNAME
            // 
            this.ITEMNAME.DataPropertyName = "ITEMNAME";
            this.ITEMNAME.HeaderText = "品名規格";
            this.ITEMNAME.Name = "ITEMNAME";
            this.ITEMNAME.Width = 78;
            // 
            // UNIT
            // 
            this.UNIT.DataPropertyName = "UNIT";
            this.UNIT.HeaderText = "單位";
            this.UNIT.Name = "UNIT";
            this.UNIT.Width = 54;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "ENABLE";
            this.dataGridViewTextBoxColumn5.HeaderText = "隱藏";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTextBoxColumn5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewTextBoxColumn5.Width = 54;
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
            this.tableAdapterManager.GB_FPRODUCTTableAdapter = this.gB_FPRODUCTTableAdapter;
            this.tableAdapterManager.GB_FRIEND1TableAdapter = null;
            this.tableAdapterManager.GB_FRIENDTableAdapter = null;
            this.tableAdapterManager.GB_FRIGHTTableAdapter = null;
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
            // GB_FPRODUCT
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(960, 624);
            this.Controls.Add(this.gB_FPRODUCTDataGridView);
            this.Controls.Add(this.gB_FPRODUCTBindingNavigator);
            this.Name = "GB_FPRODUCT";
            this.Text = "GB_FPRODUCT";
            this.Load += new System.EventHandler(this.GB_FPRODUCT_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pOTATO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FPRODUCTBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FPRODUCTBindingNavigator)).EndInit();
            this.gB_FPRODUCTBindingNavigator.ResumeLayout(false);
            this.gB_FPRODUCTBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_FPRODUCTDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.POTATO pOTATO;
        private System.Windows.Forms.BindingSource gB_FPRODUCTBindingSource;
        private ACMEDataSet.POTATOTableAdapters.GB_FPRODUCTTableAdapter gB_FPRODUCTTableAdapter;
        private System.Windows.Forms.BindingNavigator gB_FPRODUCTBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton gB_FPRODUCTBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView gB_FPRODUCTDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn ITEMCODE;
        private System.Windows.Forms.DataGridViewTextBoxColumn ITEMNAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn UNIT;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dataGridViewTextBoxColumn5;
        private ACMEDataSet.POTATOTableAdapters.TableAdapterManager tableAdapterManager;
    }
}
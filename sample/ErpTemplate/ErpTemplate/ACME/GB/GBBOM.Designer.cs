namespace ACME
{
    partial class GBBOM
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GBBOM));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle5 = new System.Windows.Forms.DataGridViewCellStyle();
            this.gB_BOMMBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.gB_BOMMBindingSource = new System.Windows.Forms.BindingSource(this.components);
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
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.gB_BOMMBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.gB_BOMMDataGridView = new System.Windows.Forms.DataGridView();
            this.CODE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CODENAME = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PRICE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PRICED = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gB_BOMDBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gB_BOMDDataGridView = new System.Windows.Forms.DataGridView();
            this.CODE2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CODENAME2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.QTY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PTICE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AMT = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.gB_BOMMTableAdapter = new ACME.ACMEDataSet.POTATOTableAdapters.GB_BOMMTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.POTATOTableAdapters.TableAdapterManager();
            this.gB_BOMDTableAdapter = new ACME.ACMEDataSet.POTATOTableAdapters.GB_BOMDTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMMBindingNavigator)).BeginInit();
            this.gB_BOMMBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMMBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pOTATO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMMDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMDBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMDDataGridView)).BeginInit();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // gB_BOMMBindingNavigator
            // 
            this.gB_BOMMBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.gB_BOMMBindingNavigator.BindingSource = this.gB_BOMMBindingSource;
            this.gB_BOMMBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.gB_BOMMBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.gB_BOMMBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.toolStripButton1,
            this.gB_BOMMBindingNavigatorSaveItem});
            this.gB_BOMMBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.gB_BOMMBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.gB_BOMMBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.gB_BOMMBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.gB_BOMMBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.gB_BOMMBindingNavigator.Name = "gB_BOMMBindingNavigator";
            this.gB_BOMMBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.gB_BOMMBindingNavigator.Size = new System.Drawing.Size(1109, 25);
            this.gB_BOMMBindingNavigator.TabIndex = 0;
            this.gB_BOMMBindingNavigator.Text = "bindingNavigator1";
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
            // gB_BOMMBindingSource
            // 
            this.gB_BOMMBindingSource.DataMember = "GB_BOMM";
            this.gB_BOMMBindingSource.DataSource = this.pOTATO;
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
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = global::ACME.Properties.Resources.bnEdit_Image;
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton1.Text = "toolStripButton1";
            this.toolStripButton1.ToolTipText = "編輯";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // gB_BOMMBindingNavigatorSaveItem
            // 
            this.gB_BOMMBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.gB_BOMMBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("gB_BOMMBindingNavigatorSaveItem.Image")));
            this.gB_BOMMBindingNavigatorSaveItem.Name = "gB_BOMMBindingNavigatorSaveItem";
            this.gB_BOMMBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.gB_BOMMBindingNavigatorSaveItem.Text = "儲存資料";
            this.gB_BOMMBindingNavigatorSaveItem.Click += new System.EventHandler(this.gB_BOMMBindingNavigatorSaveItem_Click);
            // 
            // gB_BOMMDataGridView
            // 
            this.gB_BOMMDataGridView.AutoGenerateColumns = false;
            this.gB_BOMMDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.gB_BOMMDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gB_BOMMDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CODE,
            this.CODENAME,
            this.PRICE,
            this.PRICED});
            this.gB_BOMMDataGridView.DataSource = this.gB_BOMMBindingSource;
            this.gB_BOMMDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gB_BOMMDataGridView.Location = new System.Drawing.Point(0, 0);
            this.gB_BOMMDataGridView.Name = "gB_BOMMDataGridView";
            this.gB_BOMMDataGridView.RowTemplate.Height = 24;
            this.gB_BOMMDataGridView.Size = new System.Drawing.Size(1109, 439);
            this.gB_BOMMDataGridView.TabIndex = 1;
            this.gB_BOMMDataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.gB_BOMMDataGridView_CellValueChanged);
            // 
            // CODE
            // 
            this.CODE.DataPropertyName = "CODE";
            this.CODE.HeaderText = "母料號編碼";
            this.CODE.Name = "CODE";
            this.CODE.Width = 90;
            // 
            // CODENAME
            // 
            this.CODENAME.DataPropertyName = "CODENAME";
            this.CODENAME.HeaderText = "名稱";
            this.CODENAME.Name = "CODENAME";
            this.CODENAME.Width = 54;
            // 
            // PRICE
            // 
            this.PRICE.DataPropertyName = "PRICE";
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle1.Format = "N0";
            dataGridViewCellStyle1.NullValue = null;
            this.PRICE.DefaultCellStyle = dataGridViewCellStyle1;
            this.PRICE.HeaderText = "官網金額";
            this.PRICE.Name = "PRICE";
            this.PRICE.Width = 78;
            // 
            // PRICED
            // 
            this.PRICED.DataPropertyName = "PRICED";
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle2.Format = "N0";
            dataGridViewCellStyle2.NullValue = null;
            this.PRICED.DefaultCellStyle = dataGridViewCellStyle2;
            this.PRICED.HeaderText = "明細金額";
            this.PRICED.Name = "PRICED";
            this.PRICED.ReadOnly = true;
            this.PRICED.Width = 78;
            // 
            // gB_BOMDBindingSource
            // 
            this.gB_BOMDBindingSource.DataMember = "GB_BOMM_GB_BOMD";
            this.gB_BOMDBindingSource.DataSource = this.gB_BOMMBindingSource;
            // 
            // gB_BOMDDataGridView
            // 
            this.gB_BOMDDataGridView.AutoGenerateColumns = false;
            this.gB_BOMDDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.gB_BOMDDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gB_BOMDDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.CODE2,
            this.CODENAME2,
            this.QTY,
            this.PTICE,
            this.AMT});
            this.gB_BOMDDataGridView.DataSource = this.gB_BOMDBindingSource;
            this.gB_BOMDDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gB_BOMDDataGridView.Location = new System.Drawing.Point(0, 0);
            this.gB_BOMDDataGridView.Name = "gB_BOMDDataGridView";
            this.gB_BOMDDataGridView.RowTemplate.Height = 24;
            this.gB_BOMDDataGridView.Size = new System.Drawing.Size(1109, 211);
            this.gB_BOMDDataGridView.TabIndex = 2;
            this.gB_BOMDDataGridView.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.gB_BOMDDataGridView_CellValueChanged);
            // 
            // CODE2
            // 
            this.CODE2.DataPropertyName = "CODE";
            this.CODE2.HeaderText = "子料號編碼";
            this.CODE2.Name = "CODE2";
            this.CODE2.Width = 90;
            // 
            // CODENAME2
            // 
            this.CODENAME2.DataPropertyName = "CODENAME";
            this.CODENAME2.HeaderText = "名稱";
            this.CODENAME2.Name = "CODENAME2";
            this.CODENAME2.Width = 54;
            // 
            // QTY
            // 
            this.QTY.DataPropertyName = "QTY";
            dataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle3.Format = "N0";
            this.QTY.DefaultCellStyle = dataGridViewCellStyle3;
            this.QTY.HeaderText = "數量";
            this.QTY.Name = "QTY";
            this.QTY.Width = 54;
            // 
            // PTICE
            // 
            this.PTICE.DataPropertyName = "PTICE";
            dataGridViewCellStyle4.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle4.Format = "N1";
            dataGridViewCellStyle4.NullValue = null;
            this.PTICE.DefaultCellStyle = dataGridViewCellStyle4;
            this.PTICE.HeaderText = "單價";
            this.PTICE.Name = "PTICE";
            this.PTICE.Width = 54;
            // 
            // AMT
            // 
            this.AMT.DataPropertyName = "AMT";
            dataGridViewCellStyle5.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
            dataGridViewCellStyle5.Format = "N1";
            dataGridViewCellStyle5.NullValue = null;
            this.AMT.DefaultCellStyle = dataGridViewCellStyle5;
            this.AMT.HeaderText = "總計";
            this.AMT.Name = "AMT";
            this.AMT.ReadOnly = true;
            this.AMT.Width = 54;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.gB_BOMMDataGridView);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 25);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1109, 439);
            this.panel1.TabIndex = 3;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.gB_BOMDDataGridView);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 464);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1109, 211);
            this.panel2.TabIndex = 4;
            // 
            // gB_BOMMTableAdapter
            // 
            this.gB_BOMMTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.GB_BOMDTableAdapter = this.gB_BOMDTableAdapter;
            this.tableAdapterManager.GB_BOMMTableAdapter = this.gB_BOMMTableAdapter;
            this.tableAdapterManager.GB_CS1TableAdapter = null;
            this.tableAdapterManager.GB_CS2TableAdapter = null;
            this.tableAdapterManager.GB_CSDTableAdapter = null;
            this.tableAdapterManager.GB_CSTableAdapter = null;
            this.tableAdapterManager.GB_DATELOCKTableAdapter = null;
            this.tableAdapterManager.GB_FISHTableAdapter = null;
            this.tableAdapterManager.GB_FOC2TableAdapter = null;
            this.tableAdapterManager.GB_FOC3TableAdapter = null;
            this.tableAdapterManager.GB_FOCTableAdapter = null;
            this.tableAdapterManager.GB_FRIEND1TableAdapter = null;
            this.tableAdapterManager.GB_FRIENDTableAdapter = null;
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
            // gB_BOMDTableAdapter
            // 
            this.gB_BOMDTableAdapter.ClearBeforeFill = true;
            // 
            // GBBOM
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1109, 675);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.gB_BOMMBindingNavigator);
            this.Name = "GBBOM";
            this.Text = "產品組合";
            this.Load += new System.EventHandler(this.GBBOM_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMMBindingNavigator)).EndInit();
            this.gB_BOMMBindingNavigator.ResumeLayout(false);
            this.gB_BOMMBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMMBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pOTATO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMMDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMDBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_BOMDDataGridView)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.POTATO pOTATO;
        private System.Windows.Forms.BindingSource gB_BOMMBindingSource;
        private ACMEDataSet.POTATOTableAdapters.GB_BOMMTableAdapter gB_BOMMTableAdapter;
        private ACMEDataSet.POTATOTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.BindingNavigator gB_BOMMBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton gB_BOMMBindingNavigatorSaveItem;
        private ACMEDataSet.POTATOTableAdapters.GB_BOMDTableAdapter gB_BOMDTableAdapter;
        private System.Windows.Forms.DataGridView gB_BOMMDataGridView;
        private System.Windows.Forms.BindingSource gB_BOMDBindingSource;
        private System.Windows.Forms.DataGridView gB_BOMDDataGridView;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.DataGridViewTextBoxColumn CODE;
        private System.Windows.Forms.DataGridViewTextBoxColumn CODENAME;
        private System.Windows.Forms.DataGridViewTextBoxColumn PRICE;
        private System.Windows.Forms.DataGridViewTextBoxColumn PRICED;
        private System.Windows.Forms.DataGridViewTextBoxColumn CODE2;
        private System.Windows.Forms.DataGridViewTextBoxColumn CODENAME2;
        private System.Windows.Forms.DataGridViewTextBoxColumn QTY;
        private System.Windows.Forms.DataGridViewTextBoxColumn PTICE;
        private System.Windows.Forms.DataGridViewTextBoxColumn AMT;
    }
}
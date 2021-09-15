namespace ACME
{
    partial class KITINVOICE
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(KITINVOICE));
            this.iNVOICEDKITBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.iNVOICEDKITBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.iNVOICEDKITDataGridView = new System.Windows.Forms.DataGridView();
            this.iNVOICEDKITBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ship = new ACME.ACMEDataSet.ship();
            this.iNVOICEDKITTableAdapter = new ACME.ACMEDataSet.shipTableAdapters.INVOICEDKITTableAdapter();
            this.kITDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.QTY = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.iNVOICEDKITBindingNavigator)).BeginInit();
            this.iNVOICEDKITBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.iNVOICEDKITDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iNVOICEDKITBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).BeginInit();
            this.SuspendLayout();
            // 
            // iNVOICEDKITBindingNavigator
            // 
            this.iNVOICEDKITBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.iNVOICEDKITBindingNavigator.BindingSource = this.iNVOICEDKITBindingSource;
            this.iNVOICEDKITBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.iNVOICEDKITBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.iNVOICEDKITBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.iNVOICEDKITBindingNavigatorSaveItem});
            this.iNVOICEDKITBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.iNVOICEDKITBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.iNVOICEDKITBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.iNVOICEDKITBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.iNVOICEDKITBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.iNVOICEDKITBindingNavigator.Name = "iNVOICEDKITBindingNavigator";
            this.iNVOICEDKITBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.iNVOICEDKITBindingNavigator.Size = new System.Drawing.Size(460, 25);
            this.iNVOICEDKITBindingNavigator.TabIndex = 0;
            this.iNVOICEDKITBindingNavigator.Text = "bindingNavigator1";
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
            // iNVOICEDKITBindingNavigatorSaveItem
            // 
            this.iNVOICEDKITBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.iNVOICEDKITBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("iNVOICEDKITBindingNavigatorSaveItem.Image")));
            this.iNVOICEDKITBindingNavigatorSaveItem.Name = "iNVOICEDKITBindingNavigatorSaveItem";
            this.iNVOICEDKITBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.iNVOICEDKITBindingNavigatorSaveItem.Text = "儲存資料";
            this.iNVOICEDKITBindingNavigatorSaveItem.Click += new System.EventHandler(this.iNVOICEDKITBindingNavigatorSaveItem_Click);
            // 
            // iNVOICEDKITDataGridView
            // 
            this.iNVOICEDKITDataGridView.AutoGenerateColumns = false;
            this.iNVOICEDKITDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.iNVOICEDKITDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.kITDataGridViewTextBoxColumn,
            this.QTY});
            this.iNVOICEDKITDataGridView.DataSource = this.iNVOICEDKITBindingSource;
            this.iNVOICEDKITDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.iNVOICEDKITDataGridView.Location = new System.Drawing.Point(0, 25);
            this.iNVOICEDKITDataGridView.Name = "iNVOICEDKITDataGridView";
            this.iNVOICEDKITDataGridView.RowTemplate.Height = 24;
            this.iNVOICEDKITDataGridView.Size = new System.Drawing.Size(460, 528);
            this.iNVOICEDKITDataGridView.TabIndex = 2;
            // 
            // iNVOICEDKITBindingSource
            // 
            this.iNVOICEDKITBindingSource.DataMember = "INVOICEDKIT";
            this.iNVOICEDKITBindingSource.DataSource = this.ship;
            // 
            // ship
            // 
            this.ship.DataSetName = "ship";
            this.ship.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // iNVOICEDKITTableAdapter
            // 
            this.iNVOICEDKITTableAdapter.ClearBeforeFill = true;
            // 
            // kITDataGridViewTextBoxColumn
            // 
            this.kITDataGridViewTextBoxColumn.DataPropertyName = "KIT";
            this.kITDataGridViewTextBoxColumn.HeaderText = "子料件";
            this.kITDataGridViewTextBoxColumn.Name = "kITDataGridViewTextBoxColumn";
            this.kITDataGridViewTextBoxColumn.Width = 200;
            // 
            // QTY
            // 
            this.QTY.DataPropertyName = "QTY";
            this.QTY.HeaderText = "QTY";
            this.QTY.Name = "QTY";
            // 
            // KITINVOICE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(460, 553);
            this.Controls.Add(this.iNVOICEDKITDataGridView);
            this.Controls.Add(this.iNVOICEDKITBindingNavigator);
            this.Name = "KITINVOICE";
            this.Text = "KITINVOICE";
            this.Load += new System.EventHandler(this.KITINVOICE_Load);
            ((System.ComponentModel.ISupportInitialize)(this.iNVOICEDKITBindingNavigator)).EndInit();
            this.iNVOICEDKITBindingNavigator.ResumeLayout(false);
            this.iNVOICEDKITBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.iNVOICEDKITDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iNVOICEDKITBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ship ship;
        private System.Windows.Forms.BindingSource iNVOICEDKITBindingSource;
        private ACMEDataSet.shipTableAdapters.INVOICEDKITTableAdapter iNVOICEDKITTableAdapter;
        private System.Windows.Forms.BindingNavigator iNVOICEDKITBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton iNVOICEDKITBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView iNVOICEDKITDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn kITDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn QTY;
    }
}
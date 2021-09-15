namespace ACME
{
    partial class SHIP_FIREFEE
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SHIP_FIREFEE));
            this.ship = new ACME.ACMEDataSet.ship();
            this.sHIP_FIREFEEBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.sHIP_FIREFEETableAdapter = new ACME.ACMEDataSet.shipTableAdapters.SHIP_FIREFEETableAdapter();
            this.sHIP_FIREFEEBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.sHIP_FIREFEEBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.sHIP_FIREFEEDataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.ship)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sHIP_FIREFEEBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sHIP_FIREFEEBindingNavigator)).BeginInit();
            this.sHIP_FIREFEEBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sHIP_FIREFEEDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // ship
            // 
            this.ship.DataSetName = "ship";
            this.ship.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // sHIP_FIREFEEBindingSource
            // 
            this.sHIP_FIREFEEBindingSource.DataMember = "SHIP_FIREFEE";
            this.sHIP_FIREFEEBindingSource.DataSource = this.ship;
            // 
            // sHIP_FIREFEETableAdapter
            // 
            this.sHIP_FIREFEETableAdapter.ClearBeforeFill = true;
            // 
            // sHIP_FIREFEEBindingNavigator
            // 
            this.sHIP_FIREFEEBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.sHIP_FIREFEEBindingNavigator.BindingSource = this.sHIP_FIREFEEBindingSource;
            this.sHIP_FIREFEEBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.sHIP_FIREFEEBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.sHIP_FIREFEEBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.sHIP_FIREFEEBindingNavigatorSaveItem});
            this.sHIP_FIREFEEBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.sHIP_FIREFEEBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.sHIP_FIREFEEBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.sHIP_FIREFEEBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.sHIP_FIREFEEBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.sHIP_FIREFEEBindingNavigator.Name = "sHIP_FIREFEEBindingNavigator";
            this.sHIP_FIREFEEBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.sHIP_FIREFEEBindingNavigator.Size = new System.Drawing.Size(474, 25);
            this.sHIP_FIREFEEBindingNavigator.TabIndex = 0;
            this.sHIP_FIREFEEBindingNavigator.Text = "bindingNavigator1";
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
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(27, 22);
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
            // sHIP_FIREFEEBindingNavigatorSaveItem
            // 
            this.sHIP_FIREFEEBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.sHIP_FIREFEEBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("sHIP_FIREFEEBindingNavigatorSaveItem.Image")));
            this.sHIP_FIREFEEBindingNavigatorSaveItem.Name = "sHIP_FIREFEEBindingNavigatorSaveItem";
            this.sHIP_FIREFEEBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.sHIP_FIREFEEBindingNavigatorSaveItem.Text = "儲存資料";
            this.sHIP_FIREFEEBindingNavigatorSaveItem.Click += new System.EventHandler(this.sHIP_FIREFEEBindingNavigatorSaveItem_Click);
            // 
            // sHIP_FIREFEEDataGridView
            // 
            this.sHIP_FIREFEEDataGridView.AutoGenerateColumns = false;
            this.sHIP_FIREFEEDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.sHIP_FIREFEEDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3});
            this.sHIP_FIREFEEDataGridView.DataSource = this.sHIP_FIREFEEBindingSource;
            this.sHIP_FIREFEEDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sHIP_FIREFEEDataGridView.Location = new System.Drawing.Point(0, 25);
            this.sHIP_FIREFEEDataGridView.Name = "sHIP_FIREFEEDataGridView";
            this.sHIP_FIREFEEDataGridView.RowTemplate.Height = 24;
            this.sHIP_FIREFEEDataGridView.Size = new System.Drawing.Size(474, 390);
            this.sHIP_FIREFEEDataGridView.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "ID";
            this.dataGridViewTextBoxColumn1.HeaderText = "ID";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Visible = false;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "DOCDATE";
            this.dataGridViewTextBoxColumn2.HeaderText = "年月";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "FEE";
            this.dataGridViewTextBoxColumn3.HeaderText = "費率";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // SHIP_FIREFEE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(474, 415);
            this.Controls.Add(this.sHIP_FIREFEEDataGridView);
            this.Controls.Add(this.sHIP_FIREFEEBindingNavigator);
            this.Name = "SHIP_FIREFEE";
            this.Text = "燃油費率";
            this.Load += new System.EventHandler(this.SHIP_FIREFEE_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ship)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sHIP_FIREFEEBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sHIP_FIREFEEBindingNavigator)).EndInit();
            this.sHIP_FIREFEEBindingNavigator.ResumeLayout(false);
            this.sHIP_FIREFEEBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sHIP_FIREFEEDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ship ship;
        private System.Windows.Forms.BindingSource sHIP_FIREFEEBindingSource;
        private ACMEDataSet.shipTableAdapters.SHIP_FIREFEETableAdapter sHIP_FIREFEETableAdapter;
        private System.Windows.Forms.BindingNavigator sHIP_FIREFEEBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton sHIP_FIREFEEBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView sHIP_FIREFEEDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
    }
}
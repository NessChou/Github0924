namespace ACME
{
    partial class WH_ITEM1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WH_ITEM1));
            this.wh = new ACME.ACMEDataSet.wh();
            this.wH_ITM1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.wH_ITM1TableAdapter = new ACME.ACMEDataSet.whTableAdapters.WH_ITM1TableAdapter();
            this.wH_ITM1BindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.wH_ITM1BindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.wH_ITM1DataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.wh)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.wH_ITM1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.wH_ITM1BindingNavigator)).BeginInit();
            this.wH_ITM1BindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.wH_ITM1DataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // wh
            // 
            this.wh.DataSetName = "wh";
            this.wh.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // wH_ITM1BindingSource
            // 
            this.wH_ITM1BindingSource.DataMember = "WH_ITM1";
            this.wH_ITM1BindingSource.DataSource = this.wh;
            // 
            // wH_ITM1TableAdapter
            // 
            this.wH_ITM1TableAdapter.ClearBeforeFill = true;
            // 
            // wH_ITM1BindingNavigator
            // 
            this.wH_ITM1BindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.wH_ITM1BindingNavigator.BindingSource = this.wH_ITM1BindingSource;
            this.wH_ITM1BindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.wH_ITM1BindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.wH_ITM1BindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.wH_ITM1BindingNavigatorSaveItem});
            this.wH_ITM1BindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.wH_ITM1BindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.wH_ITM1BindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.wH_ITM1BindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.wH_ITM1BindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.wH_ITM1BindingNavigator.Name = "wH_ITM1BindingNavigator";
            this.wH_ITM1BindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.wH_ITM1BindingNavigator.Size = new System.Drawing.Size(968, 25);
            this.wH_ITM1BindingNavigator.TabIndex = 0;
            this.wH_ITM1BindingNavigator.Text = "bindingNavigator1";
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
            // wH_ITM1BindingNavigatorSaveItem
            // 
            this.wH_ITM1BindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.wH_ITM1BindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("wH_ITM1BindingNavigatorSaveItem.Image")));
            this.wH_ITM1BindingNavigatorSaveItem.Name = "wH_ITM1BindingNavigatorSaveItem";
            this.wH_ITM1BindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.wH_ITM1BindingNavigatorSaveItem.Text = "儲存資料";
            this.wH_ITM1BindingNavigatorSaveItem.Click += new System.EventHandler(this.wH_ITM1BindingNavigatorSaveItem_Click);
            // 
            // wH_ITM1DataGridView
            // 
            this.wH_ITM1DataGridView.AutoGenerateColumns = false;
            this.wH_ITM1DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.wH_ITM1DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.wH_ITM1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3});
            this.wH_ITM1DataGridView.DataSource = this.wH_ITM1BindingSource;
            this.wH_ITM1DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wH_ITM1DataGridView.Location = new System.Drawing.Point(0, 25);
            this.wH_ITM1DataGridView.Name = "wH_ITM1DataGridView";
            this.wH_ITM1DataGridView.RowTemplate.Height = 24;
            this.wH_ITM1DataGridView.Size = new System.Drawing.Size(968, 552);
            this.wH_ITM1DataGridView.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "Value1";
            this.dataGridViewTextBoxColumn2.HeaderText = "料號群組";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 78;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "ITEM1";
            this.dataGridViewTextBoxColumn3.HeaderText = "群組次分類";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 90;
            // 
            // WH_ITEM1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(968, 577);
            this.Controls.Add(this.wH_ITM1DataGridView);
            this.Controls.Add(this.wH_ITM1BindingNavigator);
            this.Name = "WH_ITEM1";
            this.Text = "SAP料號群組維護";
            this.Load += new System.EventHandler(this.WH_ITEM1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.wh)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.wH_ITM1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.wH_ITM1BindingNavigator)).EndInit();
            this.wH_ITM1BindingNavigator.ResumeLayout(false);
            this.wH_ITM1BindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.wH_ITM1DataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.wh wh;
        private System.Windows.Forms.BindingSource wH_ITM1BindingSource;
        private ACMEDataSet.whTableAdapters.WH_ITM1TableAdapter wH_ITM1TableAdapter;
        private System.Windows.Forms.BindingNavigator wH_ITM1BindingNavigator;
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
        private System.Windows.Forms.ToolStripButton wH_ITM1BindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView wH_ITM1DataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
    }
}
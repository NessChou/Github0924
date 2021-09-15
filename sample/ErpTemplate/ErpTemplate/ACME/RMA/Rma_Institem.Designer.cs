namespace ACME
{
    partial class Rma_Institem
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Rma_Institem));
            this.rm = new ACME.ACMEDataSet.rm();
            this.rma_InsuBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.rma_InsuTableAdapter = new ACME.ACMEDataSet.rmTableAdapters.Rma_InsuTableAdapter();
            this.rma_InsuBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.rma_InsuBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.rma_InsuDataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.rma_Insu1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.rma_Insu1TableAdapter = new ACME.ACMEDataSet.rmTableAdapters.Rma_Insu1TableAdapter();
            this.rma_Insu1DataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.rm)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rma_InsuBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rma_InsuBindingNavigator)).BeginInit();
            this.rma_InsuBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rma_InsuDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rma_Insu1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.rma_Insu1DataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // rm
            // 
            this.rm.DataSetName = "rm";
            this.rm.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // rma_InsuBindingSource
            // 
            this.rma_InsuBindingSource.DataMember = "Rma_Insu";
            this.rma_InsuBindingSource.DataSource = this.rm;
            // 
            // rma_InsuTableAdapter
            // 
            this.rma_InsuTableAdapter.ClearBeforeFill = true;
            // 
            // rma_InsuBindingNavigator
            // 
            this.rma_InsuBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.rma_InsuBindingNavigator.BindingSource = this.rma_InsuBindingSource;
            this.rma_InsuBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.rma_InsuBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.rma_InsuBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.rma_InsuBindingNavigatorSaveItem});
            this.rma_InsuBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.rma_InsuBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.rma_InsuBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.rma_InsuBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.rma_InsuBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.rma_InsuBindingNavigator.Name = "rma_InsuBindingNavigator";
            this.rma_InsuBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.rma_InsuBindingNavigator.Size = new System.Drawing.Size(505, 25);
            this.rma_InsuBindingNavigator.TabIndex = 0;
            this.rma_InsuBindingNavigator.Text = "bindingNavigator1";
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
            this.bindingNavigatorPositionItem.Size = new System.Drawing.Size(50, 22);
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
            // rma_InsuBindingNavigatorSaveItem
            // 
            this.rma_InsuBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.rma_InsuBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("rma_InsuBindingNavigatorSaveItem.Image")));
            this.rma_InsuBindingNavigatorSaveItem.Name = "rma_InsuBindingNavigatorSaveItem";
            this.rma_InsuBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.rma_InsuBindingNavigatorSaveItem.Text = "儲存資料";
            this.rma_InsuBindingNavigatorSaveItem.Click += new System.EventHandler(this.rma_InsuBindingNavigatorSaveItem_Click_1);
            // 
            // rma_InsuDataGridView
            // 
            this.rma_InsuDataGridView.AutoGenerateColumns = false;
            this.rma_InsuDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1});
            this.rma_InsuDataGridView.DataSource = this.rma_InsuBindingSource;
            this.rma_InsuDataGridView.Location = new System.Drawing.Point(12, 39);
            this.rma_InsuDataGridView.Name = "rma_InsuDataGridView";
            this.rma_InsuDataGridView.RowTemplate.Height = 24;
            this.rma_InsuDataGridView.Size = new System.Drawing.Size(145, 434);
            this.rma_InsuDataGridView.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "Category";
            this.dataGridViewTextBoxColumn1.HeaderText = "分類";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            // 
            // rma_Insu1BindingSource
            // 
            this.rma_Insu1BindingSource.DataMember = "Rma_Insu_Rma_Insu1";
            this.rma_Insu1BindingSource.DataSource = this.rma_InsuBindingSource;
            // 
            // rma_Insu1TableAdapter
            // 
            this.rma_Insu1TableAdapter.ClearBeforeFill = true;
            // 
            // rma_Insu1DataGridView
            // 
            this.rma_Insu1DataGridView.AutoGenerateColumns = false;
            this.rma_Insu1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn4});
            this.rma_Insu1DataGridView.DataSource = this.rma_Insu1BindingSource;
            this.rma_Insu1DataGridView.Location = new System.Drawing.Point(254, 39);
            this.rma_Insu1DataGridView.Name = "rma_Insu1DataGridView";
            this.rma_Insu1DataGridView.RowTemplate.Height = 24;
            this.rma_Insu1DataGridView.Size = new System.Drawing.Size(153, 434);
            this.rma_Insu1DataGridView.TabIndex = 2;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "itemcode";
            this.dataGridViewTextBoxColumn4.HeaderText = "料號";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // Rma_Institem
            // 
            this.ClientSize = new System.Drawing.Size(505, 704);
            this.Controls.Add(this.rma_Insu1DataGridView);
            this.Controls.Add(this.rma_InsuDataGridView);
            this.Controls.Add(this.rma_InsuBindingNavigator);
            this.Name = "Rma_Institem";
            this.Load += new System.EventHandler(this.Rma_Institem_Load);
            ((System.ComponentModel.ISupportInitialize)(this.rm)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rma_InsuBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rma_InsuBindingNavigator)).EndInit();
            this.rma_InsuBindingNavigator.ResumeLayout(false);
            this.rma_InsuBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rma_InsuDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rma_Insu1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.rma_Insu1DataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.rm rm;
        private System.Windows.Forms.BindingSource rma_InsuBindingSource;
        private ACME.ACMEDataSet.rmTableAdapters.Rma_InsuTableAdapter rma_InsuTableAdapter;
        private System.Windows.Forms.BindingNavigator rma_InsuBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton rma_InsuBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView rma_InsuDataGridView;
        private System.Windows.Forms.BindingSource rma_Insu1BindingSource;
        private ACME.ACMEDataSet.rmTableAdapters.Rma_Insu1TableAdapter rma_Insu1TableAdapter;
        private System.Windows.Forms.DataGridView rma_Insu1DataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;

    }
}
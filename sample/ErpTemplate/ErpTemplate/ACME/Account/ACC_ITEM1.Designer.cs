namespace ACME
{
    partial class ACC_ITEM1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ACC_ITEM1));
            this.accBank = new ACME.ACMEDataSet.AccBank();
            this.account_ITEM1BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.account_ITEM1TableAdapter = new ACME.ACMEDataSet.AccBankTableAdapters.Account_ITEM1TableAdapter();
            this.account_ITEM1BindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorDeleteItem = new System.Windows.Forms.ToolStripButton();
            this.account_ITEM1BindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.account_ITEM1DataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.accBank)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.account_ITEM1BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.account_ITEM1BindingNavigator)).BeginInit();
            this.account_ITEM1BindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.account_ITEM1DataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // accBank
            // 
            this.accBank.DataSetName = "AccBank";
            this.accBank.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // account_ITEM1BindingSource
            // 
            this.account_ITEM1BindingSource.DataMember = "Account_ITEM1";
            this.account_ITEM1BindingSource.DataSource = this.accBank;
            // 
            // account_ITEM1TableAdapter
            // 
            this.account_ITEM1TableAdapter.ClearBeforeFill = true;
            // 
            // account_ITEM1BindingNavigator
            // 
            this.account_ITEM1BindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.account_ITEM1BindingNavigator.BindingSource = this.account_ITEM1BindingSource;
            this.account_ITEM1BindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.account_ITEM1BindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.account_ITEM1BindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.account_ITEM1BindingNavigatorSaveItem});
            this.account_ITEM1BindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.account_ITEM1BindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.account_ITEM1BindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.account_ITEM1BindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.account_ITEM1BindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.account_ITEM1BindingNavigator.Name = "account_ITEM1BindingNavigator";
            this.account_ITEM1BindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.account_ITEM1BindingNavigator.Size = new System.Drawing.Size(642, 25);
            this.account_ITEM1BindingNavigator.TabIndex = 0;
            this.account_ITEM1BindingNavigator.Text = "bindingNavigator1";
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
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(28, 16);
            this.bindingNavigatorCountItem.Text = "/{0}";
            this.bindingNavigatorCountItem.ToolTipText = "項目總數";
            // 
            // bindingNavigatorSeparator1
            // 
            this.bindingNavigatorSeparator1.Name = "bindingNavigatorSeparator";
            this.bindingNavigatorSeparator1.Size = new System.Drawing.Size(6, 6);
            // 
            // bindingNavigatorMoveNextItem
            // 
            this.bindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveNextItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveNextItem.Image")));
            this.bindingNavigatorMoveNextItem.Name = "bindingNavigatorMoveNextItem";
            this.bindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveNextItem.Size = new System.Drawing.Size(23, 20);
            this.bindingNavigatorMoveNextItem.Text = "移到下一個";
            // 
            // bindingNavigatorMoveLastItem
            // 
            this.bindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorMoveLastItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorMoveLastItem.Image")));
            this.bindingNavigatorMoveLastItem.Name = "bindingNavigatorMoveLastItem";
            this.bindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorMoveLastItem.Size = new System.Drawing.Size(23, 20);
            this.bindingNavigatorMoveLastItem.Text = "移到最後面";
            // 
            // bindingNavigatorSeparator2
            // 
            this.bindingNavigatorSeparator2.Name = "bindingNavigatorSeparator";
            this.bindingNavigatorSeparator2.Size = new System.Drawing.Size(6, 6);
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
            // bindingNavigatorDeleteItem
            // 
            this.bindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.bindingNavigatorDeleteItem.Image = ((System.Drawing.Image)(resources.GetObject("bindingNavigatorDeleteItem.Image")));
            this.bindingNavigatorDeleteItem.Name = "bindingNavigatorDeleteItem";
            this.bindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = true;
            this.bindingNavigatorDeleteItem.Size = new System.Drawing.Size(23, 20);
            this.bindingNavigatorDeleteItem.Text = "刪除";
            // 
            // account_ITEM1BindingNavigatorSaveItem
            // 
            this.account_ITEM1BindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.account_ITEM1BindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("account_ITEM1BindingNavigatorSaveItem.Image")));
            this.account_ITEM1BindingNavigatorSaveItem.Name = "account_ITEM1BindingNavigatorSaveItem";
            this.account_ITEM1BindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 23);
            this.account_ITEM1BindingNavigatorSaveItem.Text = "儲存資料";
            this.account_ITEM1BindingNavigatorSaveItem.Click += new System.EventHandler(this.account_ITEM1BindingNavigatorSaveItem_Click);
            // 
            // account_ITEM1DataGridView
            // 
            this.account_ITEM1DataGridView.AutoGenerateColumns = false;
            this.account_ITEM1DataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.account_ITEM1DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.account_ITEM1DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn2});
            this.account_ITEM1DataGridView.DataSource = this.account_ITEM1BindingSource;
            this.account_ITEM1DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.account_ITEM1DataGridView.Location = new System.Drawing.Point(0, 25);
            this.account_ITEM1DataGridView.Name = "account_ITEM1DataGridView";
            this.account_ITEM1DataGridView.RowTemplate.Height = 24;
            this.account_ITEM1DataGridView.Size = new System.Drawing.Size(642, 463);
            this.account_ITEM1DataGridView.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "Value1";
            this.dataGridViewTextBoxColumn2.HeaderText = "付款條件";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 78;
            // 
            // ACC_ITEM1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(642, 488);
            this.Controls.Add(this.account_ITEM1DataGridView);
            this.Controls.Add(this.account_ITEM1BindingNavigator);
            this.Name = "ACC_ITEM1";
            this.Text = "付款條件";
            this.Load += new System.EventHandler(this.ACC_ITEM1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.accBank)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.account_ITEM1BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.account_ITEM1BindingNavigator)).EndInit();
            this.account_ITEM1BindingNavigator.ResumeLayout(false);
            this.account_ITEM1BindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.account_ITEM1DataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.AccBank accBank;
        private System.Windows.Forms.BindingSource account_ITEM1BindingSource;
        private ACMEDataSet.AccBankTableAdapters.Account_ITEM1TableAdapter account_ITEM1TableAdapter;
        private System.Windows.Forms.BindingNavigator account_ITEM1BindingNavigator;
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
        private System.Windows.Forms.ToolStripButton account_ITEM1BindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView account_ITEM1DataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
    }
}
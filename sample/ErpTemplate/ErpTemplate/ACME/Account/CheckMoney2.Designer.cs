namespace ACME
{
    partial class CheckMoney2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CheckMoney2));
            this.account_JEBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.account_JEBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.accBank = new ACME.ACMEDataSet.AccBank();
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.account_JEBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.account_JEDataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.account_JETableAdapter = new ACME.ACMEDataSet.AccBankTableAdapters.Account_JETableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.account_JEBindingNavigator)).BeginInit();
            this.account_JEBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.account_JEBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.accBank)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.account_JEDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // account_JEBindingNavigator
            // 
            this.account_JEBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.account_JEBindingNavigator.BindingSource = this.account_JEBindingSource;
            this.account_JEBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.account_JEBindingNavigator.DeleteItem = null;
            this.account_JEBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.account_JEBindingNavigatorSaveItem});
            this.account_JEBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.account_JEBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.account_JEBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.account_JEBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.account_JEBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.account_JEBindingNavigator.Name = "account_JEBindingNavigator";
            this.account_JEBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.account_JEBindingNavigator.Size = new System.Drawing.Size(1031, 25);
            this.account_JEBindingNavigator.TabIndex = 0;
            this.account_JEBindingNavigator.Text = "bindingNavigator1";
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
            // account_JEBindingSource
            // 
            this.account_JEBindingSource.DataMember = "Account_JE";
            this.account_JEBindingSource.DataSource = this.accBank;
            // 
            // accBank
            // 
            this.accBank.DataSetName = "AccBank";
            this.accBank.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(28, 22);
            this.bindingNavigatorCountItem.Text = "/{0}";
            this.bindingNavigatorCountItem.ToolTipText = "項目總數";
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
            // account_JEBindingNavigatorSaveItem
            // 
            this.account_JEBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.account_JEBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("account_JEBindingNavigatorSaveItem.Image")));
            this.account_JEBindingNavigatorSaveItem.Name = "account_JEBindingNavigatorSaveItem";
            this.account_JEBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.account_JEBindingNavigatorSaveItem.Text = "儲存資料";
            this.account_JEBindingNavigatorSaveItem.Click += new System.EventHandler(this.account_JEBindingNavigatorSaveItem_Click);
            // 
            // account_JEDataGridView
            // 
            this.account_JEDataGridView.AutoGenerateColumns = false;
            this.account_JEDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2});
            this.account_JEDataGridView.DataSource = this.account_JEBindingSource;
            this.account_JEDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.account_JEDataGridView.Location = new System.Drawing.Point(0, 25);
            this.account_JEDataGridView.Name = "account_JEDataGridView";
            this.account_JEDataGridView.RowTemplate.Height = 24;
            this.account_JEDataGridView.Size = new System.Drawing.Size(1031, 562);
            this.account_JEDataGridView.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "ID";
            this.dataGridViewTextBoxColumn1.HeaderText = "ID";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "JE";
            this.dataGridViewTextBoxColumn2.HeaderText = "傳票號碼";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            // 
            // account_JETableAdapter
            // 
            this.account_JETableAdapter.ClearBeforeFill = true;
            // 
            // CheckMoney2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1031, 587);
            this.Controls.Add(this.account_JEDataGridView);
            this.Controls.Add(this.account_JEBindingNavigator);
            this.Name = "CheckMoney2";
            this.Text = "傳票號碼";
            this.Load += new System.EventHandler(this.CheckMoney2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.account_JEBindingNavigator)).EndInit();
            this.account_JEBindingNavigator.ResumeLayout(false);
            this.account_JEBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.account_JEBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.accBank)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.account_JEDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACME.ACMEDataSet.AccBank accBank;
        private System.Windows.Forms.BindingSource account_JEBindingSource;
        private ACME.ACMEDataSet.AccBankTableAdapters.Account_JETableAdapter account_JETableAdapter;
        private System.Windows.Forms.BindingNavigator account_JEBindingNavigator;
        private System.Windows.Forms.ToolStripButton bindingNavigatorAddNewItem;
        private System.Windows.Forms.ToolStripLabel bindingNavigatorCountItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveFirstItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMovePreviousItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator;
        private System.Windows.Forms.ToolStripTextBox bindingNavigatorPositionItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator1;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveNextItem;
        private System.Windows.Forms.ToolStripButton bindingNavigatorMoveLastItem;
        private System.Windows.Forms.ToolStripSeparator bindingNavigatorSeparator2;
        private System.Windows.Forms.ToolStripButton account_JEBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView account_JEDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
    }
}
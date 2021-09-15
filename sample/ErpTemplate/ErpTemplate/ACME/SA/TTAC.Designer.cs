namespace ACME
{
    partial class TTAC
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TTAC));
            this.sa = new ACME.ACMEDataSet.sa();
            this.sATT_ACCBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.sATT_ACCTableAdapter = new ACME.ACMEDataSet.saTableAdapters.SATT_ACCTableAdapter();
            this.sATT_ACCBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.sATT_ACCBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.sATT_ACCDataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.PAY = new System.Windows.Forms.DataGridViewComboBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewComboBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.sa)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sATT_ACCBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sATT_ACCBindingNavigator)).BeginInit();
            this.sATT_ACCBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sATT_ACCDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // sa
            // 
            this.sa.DataSetName = "sa";
            this.sa.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // sATT_ACCBindingSource
            // 
            this.sATT_ACCBindingSource.DataMember = "SATT_ACC";
            this.sATT_ACCBindingSource.DataSource = this.sa;
            // 
            // sATT_ACCTableAdapter
            // 
            this.sATT_ACCTableAdapter.ClearBeforeFill = true;
            // 
            // sATT_ACCBindingNavigator
            // 
            this.sATT_ACCBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.sATT_ACCBindingNavigator.BindingSource = this.sATT_ACCBindingSource;
            this.sATT_ACCBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.sATT_ACCBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.sATT_ACCBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.sATT_ACCBindingNavigatorSaveItem});
            this.sATT_ACCBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.sATT_ACCBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.sATT_ACCBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.sATT_ACCBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.sATT_ACCBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.sATT_ACCBindingNavigator.Name = "sATT_ACCBindingNavigator";
            this.sATT_ACCBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.sATT_ACCBindingNavigator.Size = new System.Drawing.Size(750, 25);
            this.sATT_ACCBindingNavigator.TabIndex = 0;
            this.sATT_ACCBindingNavigator.Text = "bindingNavigator1";
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
            // sATT_ACCBindingNavigatorSaveItem
            // 
            this.sATT_ACCBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.sATT_ACCBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("sATT_ACCBindingNavigatorSaveItem.Image")));
            this.sATT_ACCBindingNavigatorSaveItem.Name = "sATT_ACCBindingNavigatorSaveItem";
            this.sATT_ACCBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.sATT_ACCBindingNavigatorSaveItem.Text = "儲存資料";
            this.sATT_ACCBindingNavigatorSaveItem.Click += new System.EventHandler(this.sATT_ACCBindingNavigatorSaveItem_Click);
            // 
            // sATT_ACCDataGridView
            // 
            this.sATT_ACCDataGridView.AutoGenerateColumns = false;
            this.sATT_ACCDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.sATT_ACCDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.sATT_ACCDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.PAY,
            this.dataGridViewTextBoxColumn5});
            this.sATT_ACCDataGridView.DataSource = this.sATT_ACCBindingSource;
            this.sATT_ACCDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.sATT_ACCDataGridView.Location = new System.Drawing.Point(0, 25);
            this.sATT_ACCDataGridView.Name = "sATT_ACCDataGridView";
            this.sATT_ACCDataGridView.RowTemplate.Height = 24;
            this.sATT_ACCDataGridView.Size = new System.Drawing.Size(750, 548);
            this.sATT_ACCDataGridView.TabIndex = 1;
            this.sATT_ACCDataGridView.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.sATT_ACCDataGridView_DataError);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.DataPropertyName = "ID";
            this.dataGridViewTextBoxColumn1.HeaderText = "ID";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 42;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "ACCCODE";
            this.dataGridViewTextBoxColumn2.HeaderText = "會計科目";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 78;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "ACCNAME";
            this.dataGridViewTextBoxColumn3.HeaderText = "會計科目名稱";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 72;
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "BANK";
            this.dataGridViewTextBoxColumn4.HeaderText = "銀行";
            this.dataGridViewTextBoxColumn4.Items.AddRange(new object[] {
            "華南",
            "兆豐",
            "土銀",
            "合庫",
            "中小企銀",
            "第一(活存)",
            "第一(專戶)",
            "永豐",
            "彰化",
            "星展",
            "富邦",
            "新光",
            "台新",
            "臺灣",
            "中行-泰然",
            "中行-筍崗",
            "國泰",
            "中國-招行",
            "其他"});
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTextBoxColumn4.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewTextBoxColumn4.Width = 51;
            // 
            // PAY
            // 
            this.PAY.DataPropertyName = "PAY";
            this.PAY.HeaderText = "付款方式";
            this.PAY.Items.AddRange(new object[] {
            "TT/LC",
            "TT",
            "LC",
            "CASH",
            "票據"});
            this.PAY.Name = "PAY";
            this.PAY.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.PAY.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.PAY.Width = 61;
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "CURRENCY";
            this.dataGridViewTextBoxColumn5.HeaderText = "幣別";
            this.dataGridViewTextBoxColumn5.Items.AddRange(new object[] {
            "USD",
            "NTD",
            "RMB"});
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTextBoxColumn5.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewTextBoxColumn5.Width = 51;
            // 
            // TTAC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(750, 573);
            this.Controls.Add(this.sATT_ACCDataGridView);
            this.Controls.Add(this.sATT_ACCBindingNavigator);
            this.Name = "TTAC";
            this.Text = "TTAC";
            this.Load += new System.EventHandler(this.TTAC_Load);
            ((System.ComponentModel.ISupportInitialize)(this.sa)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sATT_ACCBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sATT_ACCBindingNavigator)).EndInit();
            this.sATT_ACCBindingNavigator.ResumeLayout(false);
            this.sATT_ACCBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.sATT_ACCDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.sa sa;
        private System.Windows.Forms.BindingSource sATT_ACCBindingSource;
        private ACMEDataSet.saTableAdapters.SATT_ACCTableAdapter sATT_ACCTableAdapter;
        private System.Windows.Forms.BindingNavigator sATT_ACCBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton sATT_ACCBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView sATT_ACCDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewComboBoxColumn PAY;
        private System.Windows.Forms.DataGridViewComboBoxColumn dataGridViewTextBoxColumn5;
    }
}
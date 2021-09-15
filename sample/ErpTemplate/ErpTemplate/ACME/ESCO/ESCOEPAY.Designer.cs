namespace ACME
{
    partial class ESCOEPAY
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ESCOEPAY));
            this.eSCO_PAYBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.eSCO_PAYBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.eSCO_PAYDataGridView = new System.Windows.Forms.DataGridView();
            this.eSCO_PAYBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.eSCO = new ACME.ACMEDataSet.ESCO();
            this.eSCO_PAYTableAdapter = new ACME.ACMEDataSet.ESCOTableAdapters.ESCO_PAYTableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.ESCOTableAdapters.TableAdapterManager();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MEMO2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MEMO3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.eSCO_PAYBindingNavigator)).BeginInit();
            this.eSCO_PAYBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.eSCO_PAYDataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.eSCO_PAYBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.eSCO)).BeginInit();
            this.SuspendLayout();
            // 
            // eSCO_PAYBindingNavigator
            // 
            this.eSCO_PAYBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.eSCO_PAYBindingNavigator.BindingSource = this.eSCO_PAYBindingSource;
            this.eSCO_PAYBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.eSCO_PAYBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.eSCO_PAYBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.eSCO_PAYBindingNavigatorSaveItem});
            this.eSCO_PAYBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.eSCO_PAYBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.eSCO_PAYBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.eSCO_PAYBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.eSCO_PAYBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.eSCO_PAYBindingNavigator.Name = "eSCO_PAYBindingNavigator";
            this.eSCO_PAYBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.eSCO_PAYBindingNavigator.Size = new System.Drawing.Size(917, 25);
            this.eSCO_PAYBindingNavigator.TabIndex = 0;
            this.eSCO_PAYBindingNavigator.Text = "bindingNavigator1";
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
            // eSCO_PAYBindingNavigatorSaveItem
            // 
            this.eSCO_PAYBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.eSCO_PAYBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("eSCO_PAYBindingNavigatorSaveItem.Image")));
            this.eSCO_PAYBindingNavigatorSaveItem.Name = "eSCO_PAYBindingNavigatorSaveItem";
            this.eSCO_PAYBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.eSCO_PAYBindingNavigatorSaveItem.Text = "儲存資料";
            this.eSCO_PAYBindingNavigatorSaveItem.Click += new System.EventHandler(this.eSCO_PAYBindingNavigatorSaveItem_Click);
            // 
            // eSCO_PAYDataGridView
            // 
            this.eSCO_PAYDataGridView.AllowUserToDeleteRows = false;
            this.eSCO_PAYDataGridView.AutoGenerateColumns = false;
            this.eSCO_PAYDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.eSCO_PAYDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.eSCO_PAYDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn2,
            this.MEMO2,
            this.MEMO3,
            this.dataGridViewTextBoxColumn3});
            this.eSCO_PAYDataGridView.DataSource = this.eSCO_PAYBindingSource;
            this.eSCO_PAYDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.eSCO_PAYDataGridView.Location = new System.Drawing.Point(0, 25);
            this.eSCO_PAYDataGridView.Name = "eSCO_PAYDataGridView";
            this.eSCO_PAYDataGridView.RowTemplate.Height = 24;
            this.eSCO_PAYDataGridView.Size = new System.Drawing.Size(917, 392);
            this.eSCO_PAYDataGridView.TabIndex = 1;
            // 
            // eSCO_PAYBindingSource
            // 
            this.eSCO_PAYBindingSource.DataMember = "ESCO_PAY";
            this.eSCO_PAYBindingSource.DataSource = this.eSCO;
            // 
            // eSCO
            // 
            this.eSCO.DataSetName = "ESCO";
            this.eSCO.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // eSCO_PAYTableAdapter
            // 
            this.eSCO_PAYTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.ESCO_PAYTableAdapter = this.eSCO_PAYTableAdapter;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.ESCOTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "DocEntry";
            this.dataGridViewTextBoxColumn2.HeaderText = "銷售訂單單號";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 72;
            // 
            // MEMO2
            // 
            this.MEMO2.DataPropertyName = "MEMO2";
            this.MEMO2.HeaderText = "發票號碼";
            this.MEMO2.Name = "MEMO2";
            this.MEMO2.Width = 61;
            // 
            // MEMO3
            // 
            this.MEMO3.DataPropertyName = "MEMO3";
            this.MEMO3.HeaderText = "總金額";
            this.MEMO3.Name = "MEMO3";
            this.MEMO3.Width = 61;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "Enabled";
            this.dataGridViewTextBoxColumn3.FalseValue = "N";
            this.dataGridViewTextBoxColumn3.HeaderText = "啟用";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridViewTextBoxColumn3.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            this.dataGridViewTextBoxColumn3.TrueValue = "Y";
            this.dataGridViewTextBoxColumn3.Width = 51;
            // 
            // ESCOEPAY
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(917, 417);
            this.Controls.Add(this.eSCO_PAYDataGridView);
            this.Controls.Add(this.eSCO_PAYBindingNavigator);
            this.Name = "ESCOEPAY";
            this.Text = "e帳行設定";
            this.Load += new System.EventHandler(this.ESCOEPAY_Load);
            ((System.ComponentModel.ISupportInitialize)(this.eSCO_PAYBindingNavigator)).EndInit();
            this.eSCO_PAYBindingNavigator.ResumeLayout(false);
            this.eSCO_PAYBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.eSCO_PAYDataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.eSCO_PAYBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.eSCO)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.ESCO eSCO;
        private System.Windows.Forms.BindingSource eSCO_PAYBindingSource;
        private ACMEDataSet.ESCOTableAdapters.ESCO_PAYTableAdapter eSCO_PAYTableAdapter;
        private ACMEDataSet.ESCOTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.BindingNavigator eSCO_PAYBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton eSCO_PAYBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView eSCO_PAYDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn MEMO2;
        private System.Windows.Forms.DataGridViewTextBoxColumn MEMO3;
        private System.Windows.Forms.DataGridViewCheckBoxColumn dataGridViewTextBoxColumn3;
    }
}
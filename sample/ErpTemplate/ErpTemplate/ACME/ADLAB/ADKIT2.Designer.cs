namespace ACME
{
    partial class ADKIT2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADKIT2));
            this.aD_OITM2BindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
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
            this.aD_OITM2BindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.aD_OITM2DataGridView = new System.Windows.Forms.DataGridView();
            this.button1 = new System.Windows.Forms.Button();
            this.aD_OITM2BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.aD = new ACME.ACMEDataSet.AD();
            this.aD_OITM2TableAdapter = new ACME.ACMEDataSet.ADTableAdapters.AD_OITM2TableAdapter();
            this.tableAdapterManager = new ACME.ACMEDataSet.ADTableAdapters.TableAdapterManager();
            this.ID = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ITEMCODE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.check2 = new System.Windows.Forms.DataGridViewLinkColumn();
            ((System.ComponentModel.ISupportInitialize)(this.aD_OITM2BindingNavigator)).BeginInit();
            this.aD_OITM2BindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.aD_OITM2DataGridView)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aD_OITM2BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.aD)).BeginInit();
            this.SuspendLayout();
            // 
            // aD_OITM2BindingNavigator
            // 
            this.aD_OITM2BindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.aD_OITM2BindingNavigator.BindingSource = this.aD_OITM2BindingSource;
            this.aD_OITM2BindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.aD_OITM2BindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.aD_OITM2BindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.aD_OITM2BindingNavigatorSaveItem});
            this.aD_OITM2BindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.aD_OITM2BindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.aD_OITM2BindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.aD_OITM2BindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.aD_OITM2BindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.aD_OITM2BindingNavigator.Name = "aD_OITM2BindingNavigator";
            this.aD_OITM2BindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.aD_OITM2BindingNavigator.Size = new System.Drawing.Size(1306, 25);
            this.aD_OITM2BindingNavigator.TabIndex = 0;
            this.aD_OITM2BindingNavigator.Text = "bindingNavigator1";
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
            // aD_OITM2BindingNavigatorSaveItem
            // 
            this.aD_OITM2BindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.aD_OITM2BindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("aD_OITM2BindingNavigatorSaveItem.Image")));
            this.aD_OITM2BindingNavigatorSaveItem.Name = "aD_OITM2BindingNavigatorSaveItem";
            this.aD_OITM2BindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.aD_OITM2BindingNavigatorSaveItem.Text = "儲存資料";
            this.aD_OITM2BindingNavigatorSaveItem.Click += new System.EventHandler(this.aD_OITM2BindingNavigatorSaveItem_Click);
            // 
            // aD_OITM2DataGridView
            // 
            this.aD_OITM2DataGridView.AutoGenerateColumns = false;
            this.aD_OITM2DataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.aD_OITM2DataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ID,
            this.ITEMCODE,
            this.dataGridViewTextBoxColumn2,
            this.PE,
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7,
            this.check2});
            this.aD_OITM2DataGridView.DataSource = this.aD_OITM2BindingSource;
            this.aD_OITM2DataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.aD_OITM2DataGridView.Location = new System.Drawing.Point(0, 25);
            this.aD_OITM2DataGridView.Name = "aD_OITM2DataGridView";
            this.aD_OITM2DataGridView.RowTemplate.Height = 24;
            this.aD_OITM2DataGridView.Size = new System.Drawing.Size(1306, 565);
            this.aD_OITM2DataGridView.TabIndex = 2;
            this.aD_OITM2DataGridView.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.aD_OITM2DataGridView_CellContentClick);
            this.aD_OITM2DataGridView.DefaultValuesNeeded += new System.Windows.Forms.DataGridViewRowEventHandler(this.aD_OITM2DataGridView_DefaultValuesNeeded);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(282, 2);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 3;
            this.button1.Text = "附件上傳";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // aD_OITM2BindingSource
            // 
            this.aD_OITM2BindingSource.DataMember = "AD_OITM2";
            this.aD_OITM2BindingSource.DataSource = this.aD;
            // 
            // aD
            // 
            this.aD.DataSetName = "AD";
            this.aD.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // aD_OITM2TableAdapter
            // 
            this.aD_OITM2TableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.AD_OITM2TableAdapter = this.aD_OITM2TableAdapter;
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.UpdateOrder = ACME.ACMEDataSet.ADTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // ID
            // 
            this.ID.DataPropertyName = "ID";
            this.ID.HeaderText = "ID";
            this.ID.Name = "ID";
            this.ID.ReadOnly = true;
            this.ID.Visible = false;
            // 
            // ITEMCODE
            // 
            this.ITEMCODE.DataPropertyName = "ITEMCODE";
            this.ITEMCODE.HeaderText = "產品編號(料號)";
            this.ITEMCODE.Name = "ITEMCODE";
            this.ITEMCODE.ReadOnly = true;
            this.ITEMCODE.Width = 110;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "DOCDATE";
            this.dataGridViewTextBoxColumn2.HeaderText = "EC日期";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            // 
            // PE
            // 
            this.PE.DataPropertyName = "PE";
            this.PE.HeaderText = "生產批次";
            this.PE.Name = "PE";
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "CARDCODE";
            this.dataGridViewTextBoxColumn3.HeaderText = "進貨廠商名稱";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "CARDCODE2";
            this.dataGridViewTextBoxColumn4.HeaderText = "銷售客戶名稱";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "ITEMNAME";
            this.dataGridViewTextBoxColumn6.HeaderText = " EC 項目說明";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "MEMO";
            this.dataGridViewTextBoxColumn7.HeaderText = "備註摘要";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            // 
            // check2
            // 
            this.check2.DataPropertyName = "LINK";
            this.check2.HeaderText = "附件";
            this.check2.Name = "check2";
            this.check2.ReadOnly = true;
            this.check2.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.check2.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.Automatic;
            // 
            // ADKIT2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1306, 590);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.aD_OITM2DataGridView);
            this.Controls.Add(this.aD_OITM2BindingNavigator);
            this.Name = "ADKIT2";
            this.Text = "產品管理明細表";
            this.Load += new System.EventHandler(this.ADKIT2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.aD_OITM2BindingNavigator)).EndInit();
            this.aD_OITM2BindingNavigator.ResumeLayout(false);
            this.aD_OITM2BindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.aD_OITM2DataGridView)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aD_OITM2BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.aD)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACMEDataSet.AD aD;
        private System.Windows.Forms.BindingSource aD_OITM2BindingSource;
        private ACMEDataSet.ADTableAdapters.AD_OITM2TableAdapter aD_OITM2TableAdapter;
        private ACMEDataSet.ADTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.BindingNavigator aD_OITM2BindingNavigator;
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
        private System.Windows.Forms.ToolStripButton aD_OITM2BindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView aD_OITM2DataGridView;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.DataGridViewTextBoxColumn ID;
        private System.Windows.Forms.DataGridViewTextBoxColumn ITEMCODE;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn PE;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewLinkColumn check2;
    }
}
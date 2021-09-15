namespace ACME
{
    partial class POTATOAR2
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(POTATOAR2));
            this.POTATO = new ACME.ACMEDataSet.POTATO();
            this.gB_INVTRACKBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.gB_INVTRACKTableAdapter = new ACME.ACMEDataSet.POTATOTableAdapters.GB_INVTRACKTableAdapter();
            this.gB_INVTRACKBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorCountItem = new System.Windows.Forms.ToolStripLabel();
            this.bindingNavigatorMoveFirstItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMovePreviousItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorPositionItem = new System.Windows.Forms.ToolStripTextBox();
            this.bindingNavigatorSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.bindingNavigatorMoveNextItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorMoveLastItem = new System.Windows.Forms.ToolStripButton();
            this.bindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.gB_INVTRACKBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.gB_INVTRACKDataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn11 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn10 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn9 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn8 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn7 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.POTATO)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_INVTRACKBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_INVTRACKBindingNavigator)).BeginInit();
            this.gB_INVTRACKBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_INVTRACKDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // POTATO
            // 
            this.POTATO.DataSetName = "POTATO";
            this.POTATO.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // gB_INVTRACKBindingSource
            // 
            this.gB_INVTRACKBindingSource.DataMember = "GB_INVTRACK";
            this.gB_INVTRACKBindingSource.DataSource = this.POTATO;
            // 
            // gB_INVTRACKTableAdapter
            // 
            this.gB_INVTRACKTableAdapter.ClearBeforeFill = true;
            // 
            // gB_INVTRACKBindingNavigator
            // 
            this.gB_INVTRACKBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.gB_INVTRACKBindingNavigator.BindingSource = this.gB_INVTRACKBindingSource;
            this.gB_INVTRACKBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.gB_INVTRACKBindingNavigator.DeleteItem = null;
            this.gB_INVTRACKBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.gB_INVTRACKBindingNavigatorSaveItem,
            this.toolStripButton1});
            this.gB_INVTRACKBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.gB_INVTRACKBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.gB_INVTRACKBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.gB_INVTRACKBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.gB_INVTRACKBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.gB_INVTRACKBindingNavigator.Name = "gB_INVTRACKBindingNavigator";
            this.gB_INVTRACKBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.gB_INVTRACKBindingNavigator.Size = new System.Drawing.Size(1165, 25);
            this.gB_INVTRACKBindingNavigator.TabIndex = 0;
            this.gB_INVTRACKBindingNavigator.Text = "bindingNavigator1";
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
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(24, 22);
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
            // gB_INVTRACKBindingNavigatorSaveItem
            // 
            this.gB_INVTRACKBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.gB_INVTRACKBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("gB_INVTRACKBindingNavigatorSaveItem.Image")));
            this.gB_INVTRACKBindingNavigatorSaveItem.Name = "gB_INVTRACKBindingNavigatorSaveItem";
            this.gB_INVTRACKBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.gB_INVTRACKBindingNavigatorSaveItem.Text = "儲存資料";
            this.gB_INVTRACKBindingNavigatorSaveItem.Click += new System.EventHandler(this.gB_INVTRACKBindingNavigatorSaveItem_Click);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButton1.Image = global::ACME.Properties.Resources.bnSearch_Image;
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(23, 22);
            this.toolStripButton1.Text = "重新整理";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // gB_INVTRACKDataGridView
            // 
            this.gB_INVTRACKDataGridView.AutoGenerateColumns = false;
            this.gB_INVTRACKDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5,
            this.dataGridViewTextBoxColumn6,
            this.dataGridViewTextBoxColumn7,
            this.dataGridViewTextBoxColumn8,
            this.dataGridViewTextBoxColumn9,
            this.dataGridViewTextBoxColumn10,
            this.dataGridViewTextBoxColumn11});
            this.gB_INVTRACKDataGridView.DataSource = this.gB_INVTRACKBindingSource;
            this.gB_INVTRACKDataGridView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gB_INVTRACKDataGridView.Location = new System.Drawing.Point(0, 25);
            this.gB_INVTRACKDataGridView.Name = "gB_INVTRACKDataGridView";
            this.gB_INVTRACKDataGridView.RowTemplate.Height = 24;
            this.gB_INVTRACKDataGridView.Size = new System.Drawing.Size(1165, 480);
            this.gB_INVTRACKDataGridView.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn11
            // 
            this.dataGridViewTextBoxColumn11.DataPropertyName = "U_BSDT3";
            this.dataGridViewTextBoxColumn11.HeaderText = "最後開立日期";
            this.dataGridViewTextBoxColumn11.Name = "dataGridViewTextBoxColumn11";
            // 
            // dataGridViewTextBoxColumn10
            // 
            this.dataGridViewTextBoxColumn10.DataPropertyName = "U_BSRN3";
            this.dataGridViewTextBoxColumn10.HeaderText = "已用號碼";
            this.dataGridViewTextBoxColumn10.Name = "dataGridViewTextBoxColumn10";
            // 
            // dataGridViewTextBoxColumn9
            // 
            this.dataGridViewTextBoxColumn9.DataPropertyName = "U_BSRN2";
            this.dataGridViewTextBoxColumn9.HeaderText = "結束號碼";
            this.dataGridViewTextBoxColumn9.Name = "dataGridViewTextBoxColumn9";
            // 
            // dataGridViewTextBoxColumn8
            // 
            this.dataGridViewTextBoxColumn8.DataPropertyName = "U_BSRN1";
            this.dataGridViewTextBoxColumn8.HeaderText = "開始號碼";
            this.dataGridViewTextBoxColumn8.Name = "dataGridViewTextBoxColumn8";
            // 
            // dataGridViewTextBoxColumn7
            // 
            this.dataGridViewTextBoxColumn7.DataPropertyName = "U_BSTRK";
            this.dataGridViewTextBoxColumn7.HeaderText = "發票字軌";
            this.dataGridViewTextBoxColumn7.Name = "dataGridViewTextBoxColumn7";
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.DataPropertyName = "U_BSTYP";
            this.dataGridViewTextBoxColumn6.HeaderText = "發票類別";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "U_BSYEM";
            this.dataGridViewTextBoxColumn5.HeaderText = "發票年月結束";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "U_BSYNM";
            this.dataGridViewTextBoxColumn4.HeaderText = "發票年月開始";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "U_BSYMM";
            this.dataGridViewTextBoxColumn3.HeaderText = "發票申報年度月份";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.Width = 110;
            // 
            // POTATOAR2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1165, 505);
            this.Controls.Add(this.gB_INVTRACKDataGridView);
            this.Controls.Add(this.gB_INVTRACKBindingNavigator);
            this.Name = "POTATOAR2";
            this.Text = "發票字軌";
            this.Load += new System.EventHandler(this.POTATOAR2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.POTATO)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_INVTRACKBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gB_INVTRACKBindingNavigator)).EndInit();
            this.gB_INVTRACKBindingNavigator.ResumeLayout(false);
            this.gB_INVTRACKBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gB_INVTRACKDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACME.ACMEDataSet.POTATO POTATO;
        private System.Windows.Forms.BindingSource gB_INVTRACKBindingSource;
        private ACME.ACMEDataSet.POTATOTableAdapters.GB_INVTRACKTableAdapter gB_INVTRACKTableAdapter;
        private System.Windows.Forms.BindingNavigator gB_INVTRACKBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton gB_INVTRACKBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView gB_INVTRACKDataGridView;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn7;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn8;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn9;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn10;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn11;
    }
}
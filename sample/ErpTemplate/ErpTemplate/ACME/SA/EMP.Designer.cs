namespace ACME
{
    partial class EMP
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(EMP));
            this.dINBENDON_UsersBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.bindingNavigatorAddNewItem = new System.Windows.Forms.ToolStripButton();
            this.dINBENDON_UsersBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.uSERS = new ACME.ACMEDataSet.USERS();
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
            this.dINBENDON_UsersBindingNavigatorSaveItem = new System.Windows.Forms.ToolStripButton();
            this.dINBENDON_UsersDataGridView = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dINBENDON_UsersTableAdapter = new ACME.ACMEDataSet.USERSTableAdapters.DINBENDON_UsersTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_UsersBindingNavigator)).BeginInit();
            this.dINBENDON_UsersBindingNavigator.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_UsersBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uSERS)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_UsersDataGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // dINBENDON_UsersBindingNavigator
            // 
            this.dINBENDON_UsersBindingNavigator.AddNewItem = this.bindingNavigatorAddNewItem;
            this.dINBENDON_UsersBindingNavigator.BindingSource = this.dINBENDON_UsersBindingSource;
            this.dINBENDON_UsersBindingNavigator.CountItem = this.bindingNavigatorCountItem;
            this.dINBENDON_UsersBindingNavigator.DeleteItem = this.bindingNavigatorDeleteItem;
            this.dINBENDON_UsersBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
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
            this.dINBENDON_UsersBindingNavigatorSaveItem});
            this.dINBENDON_UsersBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.dINBENDON_UsersBindingNavigator.MoveFirstItem = this.bindingNavigatorMoveFirstItem;
            this.dINBENDON_UsersBindingNavigator.MoveLastItem = this.bindingNavigatorMoveLastItem;
            this.dINBENDON_UsersBindingNavigator.MoveNextItem = this.bindingNavigatorMoveNextItem;
            this.dINBENDON_UsersBindingNavigator.MovePreviousItem = this.bindingNavigatorMovePreviousItem;
            this.dINBENDON_UsersBindingNavigator.Name = "dINBENDON_UsersBindingNavigator";
            this.dINBENDON_UsersBindingNavigator.PositionItem = this.bindingNavigatorPositionItem;
            this.dINBENDON_UsersBindingNavigator.Size = new System.Drawing.Size(512, 25);
            this.dINBENDON_UsersBindingNavigator.TabIndex = 0;
            this.dINBENDON_UsersBindingNavigator.Text = "bindingNavigator1";
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
            // dINBENDON_UsersBindingSource
            // 
            this.dINBENDON_UsersBindingSource.DataMember = "DINBENDON_Users";
            this.dINBENDON_UsersBindingSource.DataSource = this.uSERS;
            // 
            // uSERS
            // 
            this.uSERS.DataSetName = "USERS";
            this.uSERS.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // bindingNavigatorCountItem
            // 
            this.bindingNavigatorCountItem.Name = "bindingNavigatorCountItem";
            this.bindingNavigatorCountItem.Size = new System.Drawing.Size(24, 22);
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
            // dINBENDON_UsersBindingNavigatorSaveItem
            // 
            this.dINBENDON_UsersBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.dINBENDON_UsersBindingNavigatorSaveItem.Image = ((System.Drawing.Image)(resources.GetObject("dINBENDON_UsersBindingNavigatorSaveItem.Image")));
            this.dINBENDON_UsersBindingNavigatorSaveItem.Name = "dINBENDON_UsersBindingNavigatorSaveItem";
            this.dINBENDON_UsersBindingNavigatorSaveItem.Size = new System.Drawing.Size(23, 22);
            this.dINBENDON_UsersBindingNavigatorSaveItem.Text = "儲存資料";
            this.dINBENDON_UsersBindingNavigatorSaveItem.Click += new System.EventHandler(this.dINBENDON_UsersBindingNavigatorSaveItem_Click);
            // 
            // dINBENDON_UsersDataGridView
            // 
            this.dINBENDON_UsersDataGridView.AutoGenerateColumns = false;
            this.dINBENDON_UsersDataGridView.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn3,
            this.dataGridViewTextBoxColumn2,
            this.dataGridViewTextBoxColumn4,
            this.dataGridViewTextBoxColumn5});
            this.dINBENDON_UsersDataGridView.DataSource = this.dINBENDON_UsersBindingSource;
            this.dINBENDON_UsersDataGridView.Location = new System.Drawing.Point(12, 39);
            this.dINBENDON_UsersDataGridView.Name = "dINBENDON_UsersDataGridView";
            this.dINBENDON_UsersDataGridView.RowTemplate.Height = 24;
            this.dINBENDON_UsersDataGridView.Size = new System.Drawing.Size(470, 581);
            this.dINBENDON_UsersDataGridView.TabIndex = 1;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.DataPropertyName = "UserName";
            this.dataGridViewTextBoxColumn3.HeaderText = "姓名";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.DataPropertyName = "UserId";
            this.dataGridViewTextBoxColumn2.HeaderText = "寄件名稱";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            // 
            // dataGridViewTextBoxColumn4
            // 
            this.dataGridViewTextBoxColumn4.DataPropertyName = "phone";
            this.dataGridViewTextBoxColumn4.HeaderText = "分機";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            // 
            // dataGridViewTextBoxColumn5
            // 
            this.dataGridViewTextBoxColumn5.DataPropertyName = "Mobile";
            this.dataGridViewTextBoxColumn5.HeaderText = "手機";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            // 
            // dINBENDON_UsersTableAdapter
            // 
            this.dINBENDON_UsersTableAdapter.ClearBeforeFill = true;
            // 
            // EMP
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(512, 621);
            this.Controls.Add(this.dINBENDON_UsersDataGridView);
            this.Controls.Add(this.dINBENDON_UsersBindingNavigator);
            this.Name = "EMP";
            this.Text = "Sales&工程式";
            this.Load += new System.EventHandler(this.EMP_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_UsersBindingNavigator)).EndInit();
            this.dINBENDON_UsersBindingNavigator.ResumeLayout(false);
            this.dINBENDON_UsersBindingNavigator.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_UsersBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uSERS)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dINBENDON_UsersDataGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private ACME.ACMEDataSet.USERS uSERS;
        private System.Windows.Forms.BindingSource dINBENDON_UsersBindingSource;
        private ACME.ACMEDataSet.USERSTableAdapters.DINBENDON_UsersTableAdapter dINBENDON_UsersTableAdapter;
        private System.Windows.Forms.BindingNavigator dINBENDON_UsersBindingNavigator;
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
        private System.Windows.Forms.ToolStripButton dINBENDON_UsersBindingNavigatorSaveItem;
        private System.Windows.Forms.DataGridView dINBENDON_UsersDataGridView;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
    }
}
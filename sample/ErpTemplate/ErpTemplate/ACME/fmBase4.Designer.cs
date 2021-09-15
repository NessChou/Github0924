namespace ACME
{
    partial class fmBase4
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(fmBase4));
            this.panel1 = new System.Windows.Forms.Panel();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.SL_Status = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.SL_RecordCount = new System.Windows.Forms.ToolStripStatusLabel();
            this.bnFirst = new System.Windows.Forms.ToolStripButton();
            this.bnPrevious = new System.Windows.Forms.ToolStripButton();
            this.bnNext = new System.Windows.Forms.ToolStripButton();
            this.bnLast = new System.Windows.Forms.ToolStripButton();
            this.BindingNavigatorSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.bnAddNew = new System.Windows.Forms.ToolStripButton();
            this.bnEdit = new System.Windows.Forms.ToolStripButton();
            this.bnDelete = new System.Windows.Forms.ToolStripButton();
            this.bnEndEdit = new System.Windows.Forms.ToolStripButton();
            this.Copy2 = new System.Windows.Forms.ToolStripButton();
            this.bnCancelEdit = new System.Windows.Forms.ToolStripButton();
            this.bnQuery = new System.Windows.Forms.ToolStripButton();
            this.bnExit = new System.Windows.Forms.ToolStripButton();
            this.ToolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.FieldComboBox = new System.Windows.Forms.ToolStripComboBox();
            this.SearchTextBox = new System.Windows.Forms.ToolStripTextBox();
            this.bnSearch = new System.Windows.Forms.ToolStripButton();
            this.bnPrint = new System.Windows.Forms.ToolStripButton();
            this.BaseBindingNavigator = new System.Windows.Forms.BindingNavigator(this.components);
            this.SAVEButton = new System.Windows.Forms.ToolStripButton();
            this.panel1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BaseBindingNavigator)).BeginInit();
            this.BaseBindingNavigator.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.statusStrip1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 38);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(834, 465);
            this.panel1.TabIndex = 2;
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.SL_Status,
            this.toolStripStatusLabel2,
            this.SL_RecordCount});
            this.statusStrip1.Location = new System.Drawing.Point(0, 443);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(834, 22);
            this.statusStrip1.TabIndex = 0;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(34, 17);
            this.toolStripStatusLabel1.Text = "狀態:";
            // 
            // SL_Status
            // 
            this.SL_Status.Name = "SL_Status";
            this.SL_Status.Size = new System.Drawing.Size(31, 17);
            this.SL_Status.Text = "瀏覽";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(49, 17);
            this.toolStripStatusLabel2.Text = "     筆數:";
            // 
            // SL_RecordCount
            // 
            this.SL_RecordCount.Name = "SL_RecordCount";
            this.SL_RecordCount.Size = new System.Drawing.Size(26, 17);
            this.SL_RecordCount.Text = "0/0";
            // 
            // bnFirst
            // 
            this.bnFirst.Image = ((System.Drawing.Image)(resources.GetObject("bnFirst.Image")));
            this.bnFirst.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnFirst.Name = "bnFirst";
            this.bnFirst.Size = new System.Drawing.Size(35, 35);
            this.bnFirst.Text = "首筆";
            this.bnFirst.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnFirst.Click += new System.EventHandler(this.bnFirst_Click);
            // 
            // bnPrevious
            // 
            this.bnPrevious.Image = ((System.Drawing.Image)(resources.GetObject("bnPrevious.Image")));
            this.bnPrevious.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnPrevious.Name = "bnPrevious";
            this.bnPrevious.Size = new System.Drawing.Size(35, 35);
            this.bnPrevious.Text = "前筆";
            this.bnPrevious.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnPrevious.Click += new System.EventHandler(this.bnPrevious_Click);
            // 
            // bnNext
            // 
            this.bnNext.Image = ((System.Drawing.Image)(resources.GetObject("bnNext.Image")));
            this.bnNext.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnNext.Name = "bnNext";
            this.bnNext.Size = new System.Drawing.Size(35, 35);
            this.bnNext.Text = "次筆";
            this.bnNext.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnNext.Click += new System.EventHandler(this.bnNext_Click);
            // 
            // bnLast
            // 
            this.bnLast.Image = ((System.Drawing.Image)(resources.GetObject("bnLast.Image")));
            this.bnLast.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnLast.Name = "bnLast";
            this.bnLast.Size = new System.Drawing.Size(35, 35);
            this.bnLast.Text = "末筆";
            this.bnLast.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnLast.Click += new System.EventHandler(this.bnLast_Click);
            // 
            // BindingNavigatorSeparator2
            // 
            this.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator2";
            this.BindingNavigatorSeparator2.Size = new System.Drawing.Size(6, 38);
            // 
            // bnAddNew
            // 
            this.bnAddNew.Image = ((System.Drawing.Image)(resources.GetObject("bnAddNew.Image")));
            this.bnAddNew.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnAddNew.Name = "bnAddNew";
            this.bnAddNew.Size = new System.Drawing.Size(35, 35);
            this.bnAddNew.Text = "新增";
            this.bnAddNew.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnAddNew.Click += new System.EventHandler(this.bnAddNew_Click);
            // 
            // bnEdit
            // 
            this.bnEdit.Image = ((System.Drawing.Image)(resources.GetObject("bnEdit.Image")));
            this.bnEdit.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnEdit.Name = "bnEdit";
            this.bnEdit.Size = new System.Drawing.Size(35, 35);
            this.bnEdit.Text = "修改";
            this.bnEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnEdit.Click += new System.EventHandler(this.bnEdit_Click);
            // 
            // bnDelete
            // 
            this.bnDelete.Image = ((System.Drawing.Image)(resources.GetObject("bnDelete.Image")));
            this.bnDelete.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.bnDelete.Name = "bnDelete";
            this.bnDelete.Size = new System.Drawing.Size(35, 35);
            this.bnDelete.Text = "刪除";
            this.bnDelete.TextDirection = System.Windows.Forms.ToolStripTextDirection.Horizontal;
            this.bnDelete.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnDelete.Click += new System.EventHandler(this.bnDelete_Click);
            // 
            // bnEndEdit
            // 
            this.bnEndEdit.Image = ((System.Drawing.Image)(resources.GetObject("bnEndEdit.Image")));
            this.bnEndEdit.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnEndEdit.Name = "bnEndEdit";
            this.bnEndEdit.Size = new System.Drawing.Size(35, 35);
            this.bnEndEdit.Text = "確認";
            this.bnEndEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnEndEdit.Click += new System.EventHandler(this.bnEndEdit_Click);
            // 
            // Copy2
            // 
            this.Copy2.Image = global::ACME.Properties.Resources.copy;
            this.Copy2.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Copy2.Name = "Copy2";
            this.Copy2.Size = new System.Drawing.Size(35, 35);
            this.Copy2.Text = "複製";
            this.Copy2.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.Copy2.Click += new System.EventHandler(this.Copy2_Click);
            // 
            // bnCancelEdit
            // 
            this.bnCancelEdit.Image = ((System.Drawing.Image)(resources.GetObject("bnCancelEdit.Image")));
            this.bnCancelEdit.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnCancelEdit.Name = "bnCancelEdit";
            this.bnCancelEdit.Size = new System.Drawing.Size(35, 35);
            this.bnCancelEdit.Text = "取消";
            this.bnCancelEdit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnCancelEdit.Click += new System.EventHandler(this.bnCancelEdit_Click);
            // 
            // bnQuery
            // 
            this.bnQuery.Image = ((System.Drawing.Image)(resources.GetObject("bnQuery.Image")));
            this.bnQuery.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.bnQuery.Name = "bnQuery";
            this.bnQuery.Size = new System.Drawing.Size(35, 35);
            this.bnQuery.Text = "查詢";
            this.bnQuery.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.bnQuery.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnQuery.Click += new System.EventHandler(this.bnQuery_Click);
            // 
            // bnExit
            // 
            this.bnExit.Image = ((System.Drawing.Image)(resources.GetObject("bnExit.Image")));
            this.bnExit.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.bnExit.Name = "bnExit";
            this.bnExit.Size = new System.Drawing.Size(35, 35);
            this.bnExit.Text = "關閉";
            this.bnExit.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.bnExit.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnExit.Click += new System.EventHandler(this.bnExit_Click);
            // 
            // ToolStripSeparator1
            // 
            this.ToolStripSeparator1.Name = "ToolStripSeparator1";
            this.ToolStripSeparator1.Size = new System.Drawing.Size(6, 38);
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(31, 35);
            this.toolStripLabel1.Text = "欄位";
            this.toolStripLabel1.Visible = false;
            // 
            // FieldComboBox
            // 
            this.FieldComboBox.Name = "FieldComboBox";
            this.FieldComboBox.Size = new System.Drawing.Size(75, 38);
            this.FieldComboBox.Visible = false;
            // 
            // SearchTextBox
            // 
            this.SearchTextBox.MaxLength = 20;
            this.SearchTextBox.Name = "SearchTextBox";
            this.SearchTextBox.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.SearchTextBox.Size = new System.Drawing.Size(140, 38);
            this.SearchTextBox.Text = " ";
            this.SearchTextBox.Visible = false;
            this.SearchTextBox.TextChanged += new System.EventHandler(this.SearchTextBox_TextChanged);
            // 
            // bnSearch
            // 
            this.bnSearch.Image = ((System.Drawing.Image)(resources.GetObject("bnSearch.Image")));
            this.bnSearch.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnSearch.Name = "bnSearch";
            this.bnSearch.Size = new System.Drawing.Size(59, 35);
            this.bnSearch.Text = "單號速查";
            this.bnSearch.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnSearch.Visible = false;
            this.bnSearch.Click += new System.EventHandler(this.bnSearch_Click);
            // 
            // bnPrint
            // 
            this.bnPrint.Image = ((System.Drawing.Image)(resources.GetObject("bnPrint.Image")));
            this.bnPrint.ImageTransparentColor = System.Drawing.Color.Olive;
            this.bnPrint.Name = "bnPrint";
            this.bnPrint.Size = new System.Drawing.Size(35, 35);
            this.bnPrint.Text = "列印";
            this.bnPrint.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.bnPrint.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.bnPrint.Visible = false;
            this.bnPrint.Click += new System.EventHandler(this.bnPrint_Click);
            // 
            // BaseBindingNavigator
            // 
            this.BaseBindingNavigator.AddNewItem = null;
            this.BaseBindingNavigator.CountItem = null;
            this.BaseBindingNavigator.DeleteItem = null;
            this.BaseBindingNavigator.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.bnFirst,
            this.bnPrevious,
            this.bnNext,
            this.bnLast,
            this.BindingNavigatorSeparator2,
            this.bnAddNew,
            this.bnEdit,
            this.bnDelete,
            this.bnEndEdit,
            this.Copy2,
            this.bnCancelEdit,
            this.bnQuery,
            this.SAVEButton,
            this.bnExit,
            this.ToolStripSeparator1,
            this.toolStripLabel1,
            this.FieldComboBox,
            this.SearchTextBox,
            this.bnSearch,
            this.bnPrint});
            this.BaseBindingNavigator.Location = new System.Drawing.Point(0, 0);
            this.BaseBindingNavigator.MoveFirstItem = null;
            this.BaseBindingNavigator.MoveLastItem = null;
            this.BaseBindingNavigator.MoveNextItem = null;
            this.BaseBindingNavigator.MovePreviousItem = null;
            this.BaseBindingNavigator.Name = "BaseBindingNavigator";
            this.BaseBindingNavigator.PositionItem = null;
            this.BaseBindingNavigator.Size = new System.Drawing.Size(834, 38);
            this.BaseBindingNavigator.TabIndex = 1;
            this.BaseBindingNavigator.Text = "BindingNavigator1";
            // 
            // SAVEButton
            // 
            this.SAVEButton.Image = ((System.Drawing.Image)(resources.GetObject("SAVEButton.Image")));
            this.SAVEButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.SAVEButton.Name = "SAVEButton";
            this.SAVEButton.Size = new System.Drawing.Size(35, 35);
            this.SAVEButton.Text = "存檔";
            this.SAVEButton.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.SAVEButton.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.SAVEButton.Visible = false;
            this.SAVEButton.Click += new System.EventHandler(this.SAVEButton_Click);
            // 
            // fmBase4
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(834, 503);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.BaseBindingNavigator);
            this.KeyPreview = true;
            this.Name = "fmBase4";
            this.Text = "fmBase";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.fmBase_FormClosing);
            this.Load += new System.EventHandler(this.fmBase_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BaseBindingNavigator)).EndInit();
            this.BaseBindingNavigator.ResumeLayout(false);
            this.BaseBindingNavigator.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel SL_Status;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripStatusLabel SL_RecordCount;
        public System.Windows.Forms.ToolStripButton bnFirst;
        public System.Windows.Forms.ToolStripButton bnPrevious;
        public System.Windows.Forms.ToolStripButton bnNext;
        public System.Windows.Forms.ToolStripButton bnLast;
        internal System.Windows.Forms.ToolStripSeparator BindingNavigatorSeparator2;
        public System.Windows.Forms.ToolStripButton bnAddNew;
        public System.Windows.Forms.ToolStripButton bnEdit;
        private System.Windows.Forms.ToolStripButton bnDelete;
        public System.Windows.Forms.ToolStripButton bnEndEdit;
        private System.Windows.Forms.ToolStripButton Copy2;
        public System.Windows.Forms.ToolStripButton bnCancelEdit;
        private System.Windows.Forms.ToolStripButton bnQuery;
        private System.Windows.Forms.ToolStripButton bnExit;
        internal System.Windows.Forms.ToolStripSeparator ToolStripSeparator1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripComboBox FieldComboBox;
        internal System.Windows.Forms.ToolStripTextBox SearchTextBox;
        internal System.Windows.Forms.ToolStripButton bnSearch;
        internal System.Windows.Forms.ToolStripButton bnPrint;
        public System.Windows.Forms.BindingNavigator BaseBindingNavigator;
        private System.Windows.Forms.ToolStripButton SAVEButton;
    }
}
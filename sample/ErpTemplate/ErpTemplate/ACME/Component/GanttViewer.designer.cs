namespace ACME
{
    partial class GanttViewer
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

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.TaskColorCategory = new System.Windows.Forms.DataGridViewImageColumn();
            this.TaskPriority = new System.Windows.Forms.DataGridViewImageColumn();
            this.Task = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.gvTitle = new System.Windows.Forms.DataGridView();
            this.TaskFlag = new System.Windows.Forms.DataGridViewImageColumn();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.gvGantt = new ACME.GanttChart();
            ((System.ComponentModel.ISupportInitialize)(this.gvTitle)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvGantt)).BeginInit();
            this.SuspendLayout();
            // 
            // TaskColorCategory
            // 
            this.TaskColorCategory.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.TaskColorCategory.DataPropertyName = "ColorCategory";
            this.TaskColorCategory.HeaderText = "";
            this.TaskColorCategory.MinimumWidth = 20;
            this.TaskColorCategory.Name = "TaskColorCategory";
            this.TaskColorCategory.Width = 20;
            // 
            // TaskPriority
            // 
            this.TaskPriority.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.TaskPriority.DataPropertyName = "Priority";
            this.TaskPriority.HeaderText = "!";
            this.TaskPriority.MinimumWidth = 20;
            this.TaskPriority.Name = "TaskPriority";
            this.TaskPriority.Width = 20;
            // 
            // Task
            // 
            this.Task.DataPropertyName = "Title";
            this.Task.HeaderText = "Task";
            this.Task.Name = "Task";
            this.Task.ReadOnly = true;
            // 
            // gvTitle
            // 
            this.gvTitle.AllowUserToAddRows = false;
            this.gvTitle.AllowUserToDeleteRows = false;
            this.gvTitle.AllowUserToResizeColumns = false;
            this.gvTitle.AllowUserToResizeRows = false;
            this.gvTitle.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.gvTitle.BackgroundColor = System.Drawing.Color.White;
            this.gvTitle.ColumnHeadersHeight = 50;
            this.gvTitle.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.gvTitle.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Task,
            this.TaskPriority,
            this.TaskColorCategory,
            this.TaskFlag});
            this.gvTitle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gvTitle.GridColor = System.Drawing.Color.Black;
            this.gvTitle.Location = new System.Drawing.Point(0, 0);
            this.gvTitle.MultiSelect = false;
            this.gvTitle.Name = "gvTitle";
            this.gvTitle.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None;
            this.gvTitle.RowHeadersVisible = false;
            this.gvTitle.RowTemplate.Height = 24;
            this.gvTitle.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gvTitle.Size = new System.Drawing.Size(203, 471);
            this.gvTitle.TabIndex = 0;
            this.gvTitle.Scroll += new System.Windows.Forms.ScrollEventHandler(this.gvTitle_Scroll);
            this.gvTitle.DoubleClick += new System.EventHandler(this.gvTitle_DoubleClick);
            this.gvTitle.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.gvTitle_CellFormatting);
            this.gvTitle.CellPainting += new System.Windows.Forms.DataGridViewCellPaintingEventHandler(this.gvTitle_CellPainting);
            // 
            // TaskFlag
            // 
            this.TaskFlag.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.None;
            this.TaskFlag.DataPropertyName = "Flag";
            this.TaskFlag.HeaderText = "";
            this.TaskFlag.MinimumWidth = 20;
            this.TaskFlag.Name = "TaskFlag";
            this.TaskFlag.Width = 20;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.gvTitle);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.gvGantt);
            this.splitContainer1.Size = new System.Drawing.Size(610, 471);
            this.splitContainer1.SplitterDistance = 203;
            this.splitContainer1.TabIndex = 1;
            // 
            // gvGantt
            // 
            this.gvGantt.AllowUserToAddRows = false;
            this.gvGantt.AllowUserToDeleteRows = false;
            this.gvGantt.AllowUserToResizeColumns = false;
            this.gvGantt.AllowUserToResizeRows = false;
            this.gvGantt.BackgroundColor = System.Drawing.Color.White;
            this.gvGantt.ColumnHeadersHeight = 50;
            this.gvGantt.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.gvGantt.Dock = System.Windows.Forms.DockStyle.Fill;
            this.gvGantt.GridColor = System.Drawing.Color.Black;
            this.gvGantt.HeaderColor = System.Drawing.Color.LightGray;
            this.gvGantt.Location = new System.Drawing.Point(0, 0);
            this.gvGantt.MultiSelect = false;
            this.gvGantt.Name = "gvGantt";
            this.gvGantt.ReadOnly = true;
            this.gvGantt.RowHeadersVisible = false;
            this.gvGantt.RowTemplate.Height = 24;
            this.gvGantt.Size = new System.Drawing.Size(403, 471);
            this.gvGantt.TabIndex = 0;
            this.gvGantt.TaskColor = System.Drawing.Color.OrangeRed;
            this.gvGantt.TodayColor = System.Drawing.Color.Green;
            this.gvGantt.WeekendColor = System.Drawing.Color.LightGray;
            this.gvGantt.Scroll += new System.Windows.Forms.ScrollEventHandler(this.gvGantt_Scroll);
            this.gvGantt.CellToolTipTextNeeded += new System.Windows.Forms.DataGridViewCellToolTipTextNeededEventHandler(this.gvGantt_CellToolTipTextNeeded);
            this.gvGantt.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.gvGantt_DataBindingComplete);
            // 
            // GanttViewer
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "GanttViewer";
            this.Size = new System.Drawing.Size(610, 471);
            ((System.ComponentModel.ISupportInitialize)(this.gvTitle)).EndInit();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gvGantt)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridViewImageColumn TaskColorCategory;
        private System.Windows.Forms.DataGridViewImageColumn TaskPriority;
        private System.Windows.Forms.DataGridViewTextBoxColumn Task;
        private System.Windows.Forms.DataGridView gvTitle;
        private System.Windows.Forms.DataGridViewImageColumn TaskFlag;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private ACME.GanttChart gvGantt;
    }
}
